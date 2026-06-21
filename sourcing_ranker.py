#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
네이버 해외직구/구매대행 소싱 후보 자동 랭킹
- 입력: 브랜드&키워드 엑셀/CSV
- 신규 발굴: 네이버 검색광고 키워드도구 연관키워드 후보 확장(선택)
- 조회: 네이버 검색광고 키워드도구(월간검색수), 네이버 쇼핑검색 API(전체/해외상품수, 국내가 샘플)
- 출력: reports/sourcing_rank_YYYYMMDD.xlsx, reports/discovered_keywords_YYYYMMDD.csv, history snapshot CSV

필수 환경변수:
  NAVER_CLIENT_ID
  NAVER_CLIENT_SECRET

선택 환경변수(월간 검색량 자동 갱신용):
  NAVER_SEARCHAD_API_KEY
  NAVER_SEARCHAD_SECRET_KEY
  NAVER_SEARCHAD_CUSTOMER_ID

사용 예:
  python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 80
  python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --discover-related --related-per-brand 5 --max-discovered 100
  python sourcing_ranker.py --input brands.csv --no-searchad
"""

from __future__ import annotations

import argparse
import base64
import csv
import dataclasses
import datetime as dt
import hashlib
import hmac
import json
import math
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests
from dotenv import load_dotenv


NAVER_OPENAPI_BASE = "https://openapi.naver.com"
NAVER_SEARCHAD_BASE = "https://api.searchad.naver.com"


# -----------------------------
# 유틸
# -----------------------------

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_keyword_series(value: Any) -> str:
    """키워드&시리즈 원문은 줄바꿈이 대표키워드 구분자라서 보존한다."""
    if value is None:
        return ""
    text = str(value)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def normalize_col_name(value: Any) -> str:
    """엑셀 헤더의 줄바꿈/공백 차이를 흡수하기 위한 정규화."""
    return re.sub(r"\s+", "", str(value or "")).lower()


def to_int(value: Any, default: int = 0) -> int:
    if value is None:
        return default
    if isinstance(value, str):
        v = value.strip().replace(",", "")
        # 검색광고 API는 검색량이 낮을 때 "< 10"처럼 줄 수 있음
        if "<" in v:
            return 0
        if v == "" or v.lower() in {"nan", "none"}:
            return default
        try:
            return int(float(v))
        except ValueError:
            return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    try:
        return int(float(value))
    except Exception:
        return default


def safe_div(n: float, d: float, default: float = 0.0) -> float:
    return default if d == 0 else n / d


def pct_rank(series: pd.Series, higher_is_better: bool = True) -> pd.Series:
    """0~100 percentile rank. 값이 동일하거나 비어도 깨지지 않게 처리."""
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    if len(s) == 0:
        return s
    ranks = s.rank(method="average", pct=True) * 100
    if not higher_is_better:
        ranks = 100 - ranks
    return ranks.clip(0, 100)


def clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


# -----------------------------
# 네이버 API 클라이언트
# -----------------------------

@dataclasses.dataclass
class NaverOpenApiClient:
    client_id: str
    client_secret: str
    timeout: int = 12

    @property
    def headers(self) -> Dict[str, str]:
        return {
            "X-Naver-Client-Id": self.client_id,
            "X-Naver-Client-Secret": self.client_secret,
        }

    def shopping_search(
        self,
        query: str,
        display: int = 10,
        sort: str = "sim",
        exclude: Optional[str] = None,
    ) -> Dict[str, Any]:
        url = f"{NAVER_OPENAPI_BASE}/v1/search/shop.json"
        params = {
            "query": query,
            "display": display,
            "start": 1,
            "sort": sort,
        }
        if exclude:
            params["exclude"] = exclude

        r = requests.get(url, headers=self.headers, params=params, timeout=self.timeout)
        if r.status_code != 200:
            raise RuntimeError(f"Shopping API error {r.status_code}: {r.text[:300]}")
        return r.json()


@dataclasses.dataclass
class NaverSearchAdClient:
    api_key: str
    secret_key: str
    customer_id: str
    timeout: int = 12

    def _signature(self, timestamp: str, method: str, uri: str) -> str:
        message = f"{timestamp}.{method}.{uri}"
        digest = hmac.new(
            self.secret_key.encode("utf-8"),
            message.encode("utf-8"),
            hashlib.sha256,
        ).digest()
        return base64.b64encode(digest).decode("utf-8")

    def _headers(self, method: str, uri: str) -> Dict[str, str]:
        timestamp = str(round(time.time() * 1000))
        return {
            "Content-Type": "application/json; charset=UTF-8",
            "X-Timestamp": timestamp,
            "X-API-KEY": self.api_key,
            "X-Customer": str(self.customer_id),
            "X-Signature": self._signature(timestamp, method, uri),
        }

    def keyword_tool(self, keyword: str, show_detail: bool = True) -> Dict[str, Any]:
        uri = "/keywordstool"
        method = "GET"
        params = {
            "hintKeywords": keyword,
            "showDetail": 1 if show_detail else 0,
        }
        r = requests.get(
            NAVER_SEARCHAD_BASE + uri,
            headers=self._headers(method, uri),
            params=params,
            timeout=self.timeout,
        )
        if r.status_code != 200:
            raise RuntimeError(f"SearchAd keywordstool error {r.status_code}: {r.text[:300]}")
        return r.json()


# -----------------------------
# 입력 처리
# -----------------------------

COLUMN_ALIASES = {
    "brand": ["브랜드명", "브랜드", "brand", "Brand"],
    "category": ["대표카테고리", "카테고리", "category", "Category"],
    "base_volume": ["검색량", "월간검색량", "search_volume", "Search Volume"],
    "keyword_series": ["키워드&시리즈", "키워드", "대표키워드", "keyword", "keywords"],
    "naver_products": ["네이버상품수", "네이버 상품수", "total_products"],
    "naver_overseas_products": ["네이버해외상품수", "네이버 해외 상품수", "overseas_products"],
}


def find_col(df: pd.DataFrame, logical: str) -> Optional[str]:
    normalized_to_original = {normalize_col_name(c): c for c in df.columns}
    for cand in COLUMN_ALIASES[logical]:
        key = normalize_col_name(cand)
        if key in normalized_to_original:
            return normalized_to_original[key]
    return None


def load_input(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(path)

    if path.suffix.lower() in [".xlsx", ".xls"]:
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)

    rename = {}
    for logical in COLUMN_ALIASES:
        col = find_col(df, logical)
        if col:
            rename[col] = logical

    df = df.rename(columns=rename)

    if "brand" not in df.columns:
        raise ValueError("입력 파일에 '브랜드명' 또는 'brand' 컬럼이 필요합니다.")

    if "keyword_series" not in df.columns:
        df["keyword_series"] = df["brand"]

    if "category" not in df.columns:
        df["category"] = ""

    if "base_volume" not in df.columns:
        df["base_volume"] = 0

    if "naver_products" not in df.columns:
        df["naver_products"] = 0

    if "naver_overseas_products" not in df.columns:
        df["naver_overseas_products"] = 0

    df = df[["brand", "category", "keyword_series", "base_volume", "naver_products", "naver_overseas_products"]].copy()
    df["brand"] = df["brand"].map(clean_text)
    df["category"] = df["category"].map(clean_text)
    # 중요: 키워드&시리즈의 줄바꿈은 대표키워드 추출에 필요하므로 clean_text를 쓰지 않는다.
    df["keyword_series"] = df["keyword_series"].map(clean_keyword_series)
    df = df[df["brand"] != ""].drop_duplicates(subset=["brand", "keyword_series"])
    df["source_type"] = "seed"
    df["discovery_seed"] = ""
    return df.reset_index(drop=True)


def pick_primary_keyword(row: pd.Series) -> str:
    raw = clean_keyword_series(row.get("keyword_series", "")) or clean_text(row.get("brand", ""))
    brand = clean_text(row.get("brand", ""))
    # 파일 안에서 줄바꿈/쉼표로 시리즈 키워드가 여러 개 들어온 경우 첫 번째를 대표로 사용.
    # 이전 버전은 줄바꿈을 먼저 공백으로 바꿔 API query가 너무 길어지는 문제가 있었다.
    parts = re.split(r"[\n\r,;/|]+", raw)
    parts = [clean_text(p) for p in parts if clean_text(p)]
    keyword = parts[0] if parts else brand
    # 방어 로직: 구분자가 없는 긴 키워드 묶음은 브랜드명으로 축소한다.
    if len(keyword) > 40 or len(keyword.encode("utf-8")) > 100 or keyword.count(" ") >= 4:
        keyword = brand
    return keyword


def api_safe_keyword(keyword: str, brand: str = "", max_chars: int = 40, max_bytes: int = 100) -> str:
    """네이버 API URL 414/검색광고 hintKeywords 400 방지용 검색어 축소."""
    keyword = clean_text(keyword)
    brand = clean_text(brand)
    if not keyword:
        return brand
    if len(keyword) > max_chars or len(keyword.encode("utf-8")) > max_bytes or keyword.count(" ") >= 4:
        return brand or keyword[:max_chars]
    return keyword


# -----------------------------
# API 조회
# -----------------------------

def parse_searchad_volume(resp: Dict[str, Any], exact_keyword: str) -> Tuple[int, Dict[str, Any]]:
    items = resp.get("keywordList") or resp.get("data") or []
    if not isinstance(items, list):
        return 0, {}

    exact_norm = re.sub(r"\s+", "", exact_keyword).lower()
    best = None

    # 1순위: 정확히 같은 relKeyword
    for item in items:
        rel = clean_text(item.get("relKeyword", ""))
        if re.sub(r"\s+", "", rel).lower() == exact_norm:
            best = item
            break

    # 2순위: 월간검색수 가장 큰 관련 키워드
    if best is None and items:
        best = max(
            items,
            key=lambda x: to_int(x.get("monthlyPcQcCnt")) + to_int(x.get("monthlyMobileQcCnt")),
        )

    if not best:
        return 0, {}

    pc = to_int(best.get("monthlyPcQcCnt"))
    mo = to_int(best.get("monthlyMobileQcCnt"))
    return pc + mo, best



def searchad_keyword_items(resp: Dict[str, Any]) -> List[Dict[str, Any]]:
    """검색광고 keywordstool 응답에서 keywordList를 안전하게 꺼낸다."""
    items = resp.get("keywordList") or resp.get("data") or []
    return items if isinstance(items, list) else []


def keyword_volume_from_item(item: Dict[str, Any]) -> int:
    return to_int(item.get("monthlyPcQcCnt")) + to_int(item.get("monthlyMobileQcCnt"))


def keyword_norm(value: Any) -> str:
    return re.sub(r"\s+", "", clean_text(value)).lower()


def is_too_broad_or_risky_discovery(keyword: str) -> bool:
    """신규 후보 자동추가에서 너무 넓거나 규제 리스크가 큰 키워드를 1차 제외."""
    k = clean_text(keyword).lower()
    if not k:
        return True
    # 너무 짧은 일반어는 쇼핑 검색이 넓게 퍼져서 소싱 후보로 부적합
    if len(k) <= 1:
        return True
    risky = [
        "영양제", "보충제", "건강식품", "다이어트", "의약품", "약", "성분", "담배",
        "전자담배", "술", "맥주", "와인", "CBD", "THC", "대마", "성인용품"
    ]
    return any(x.lower() in k for x in risky)




# -----------------------------
# 상품 단위 후보 발굴
# -----------------------------

PRODUCT_INTENT_WORDS = [
    "운동화", "러닝화", "스니커즈", "신발", "샌들", "슬리퍼", "부츠", "로퍼", "풋살화",
    "가방", "백팩", "숄더백", "토트백", "미니백", "크로스백", "파우치", "지갑", "카드지갑",
    "케이스", "폰케이스", "아이폰", "갤럭시", "맥세이프", "카드홀더",
    "양산", "우산", "우양산", "장우산", "암막양산",
    "텀블러", "보온병", "물병", "도시락", "머그", "컵", "식기",
    "수납", "랙", "트레이", "정리함", "후크", "스탠드", "홀더",
    "체어", "의자", "테이블", "캠핑", "랜턴", "머그컵",
    "키링", "키홀더", "인형", "피규어", "토이", "굿즈", "스티커",
    "볼캡", "캡", "모자", "비니", "장갑", "머플러", "양말", "벨트",
    "티셔츠", "반팔", "후드", "후드집업", "맨투맨", "가디건", "재킷", "자켓", "바지", "팬츠",
    "시계", "선글라스", "목걸이", "귀걸이", "팔찌",
    "샤프", "볼펜", "다이어리", "노트", "필통",
]

MODEL_SERIES_HINTS = [
    "에어맥스", "에어포스", "덩크", "조던", "코르테즈", "페가수스",
    "젤", "카야노", "님버스", "1130", "2160", "킨세이",
    "990", "992", "993", "2002", "530", "574", "860",
    "보스턴", "아리조나", "타스만", "클래식", "르플리아쥬", "플리아쥬",
    "보스턴백", "벨트백", "에브리웨어", "tower", "타워",
]

PRODUCT_STOPWORDS = [
    "정품", "공식", "해외직구", "일본직구", "일본", "구매대행", "무료배송",
    "당일", "배송", "세일", "할인", "추천", "인기", "신상", "남성", "여성", "공용",
    "키즈", "주니어", "병행수입", "특가", "최저가", "국내", "해외", "브랜드",
]


def compact_keyword(text: str) -> str:
    return re.sub(r"\s+", "", clean_text(text).lower())


def looks_like_product_keyword(keyword: str, brand: str = "", category: str = "") -> bool:
    """
    브랜드명 단독이 아니라 '브랜드 + 품목/모델/시리즈' 형태인지 판정한다.
    예: 나이키(False) / 나이키 운동화(True) / 나이키 에어맥스(True)
    """
    k = clean_text(keyword)
    if not k:
        return False
    if is_too_broad_or_risky_discovery(k):
        return False

    kn = compact_keyword(k)
    bn = compact_keyword(brand)
    if bn and kn == bn:
        return False

    # 너무 일반적인 단어만 있는 경우 제외
    if len(kn) <= 2:
        return False

    product_hit = any(compact_keyword(w) in kn for w in PRODUCT_INTENT_WORDS)
    model_hit = any(compact_keyword(w) in kn for w in MODEL_SERIES_HINTS)
    category_hit = bool(category) and any(compact_keyword(w) in kn for w in re.split(r"[,/ >]+", clean_text(category)) if len(compact_keyword(w)) >= 2)

    # 브랜드가 포함되어 있고, 브랜드 외에 의미 있는 토큰이 붙으면 상품 키워드로 인정
    if bn and bn in kn:
        remainder = kn.replace(bn, "", 1)
        if len(remainder) >= 2 and (product_hit or model_hit or category_hit or re.search(r"\d{2,}", remainder)):
            return True

    # 브랜드가 빠졌더라도 품목/모델 조합이면 후보로 인정. 단, 너무 광범위한 단일 품목은 제외될 수 있음.
    if (product_hit and model_hit) or (product_hit and re.search(r"\d{2,}", kn)):
        return True

    # 공백 기준 2토큰 이상이고 품목어가 있으면 후보 인정
    tokens = [t for t in re.split(r"\s+", k) if t and t not in PRODUCT_STOPWORDS]
    if len(tokens) >= 2 and product_hit:
        return True

    return False


def product_group_from_keyword(keyword: str, category: str = "") -> str:
    k = compact_keyword(keyword + " " + category)
    groups = [
        ("휴대폰/디지털소품", ["케이스", "폰케이스", "아이폰", "갤럭시", "맥세이프", "카드홀더"]),
        ("신발", ["운동화", "러닝화", "스니커즈", "신발", "샌들", "슬리퍼", "부츠", "로퍼", "풋살화"]),
        ("가방/지갑", ["가방", "백팩", "숄더백", "토트백", "미니백", "크로스백", "파우치", "지갑", "카드지갑"]),
        ("우산/양산", ["양산", "우산", "우양산", "장우산", "암막양산"]),
        ("주방/생활", ["텀블러", "보온병", "도시락", "머그", "수납", "랙", "정리함", "후크", "스탠드", "홀더"]),
        ("캠핑/아웃도어소품", ["체어", "의자", "테이블", "캠핑", "랜턴", "머그컵"]),
        ("캐릭터/굿즈", ["키링", "키홀더", "인형", "피규어", "토이", "굿즈", "스티커"]),
        ("패션소품/의류", ["볼캡", "캡", "모자", "비니", "장갑", "머플러", "양말", "티셔츠", "반팔", "후드", "가디건", "재킷", "바지", "팬츠"]),
        ("문구", ["샤프", "볼펜", "다이어리", "노트", "필통"]),
    ]
    for group, words in groups:
        if any(compact_keyword(w) in k for w in words):
            return group
    return clean_text(category) or "미분류"


def clean_shopping_title(title: str) -> str:
    text = clean_text(title)
    text = re.sub(r"\[[^\]]{0,30}\]|\([^\)]{0,30}\)", " ", text)
    text = re.sub(r"(?i)\b(무료배송|정품|공식|해외직구|일본직구|구매대행|당일발송|특가|세일)\b", " ", text)
    text = re.sub(r"[^0-9A-Za-z가-힣\-\+\s]", " ", text)
    return clean_text(text)


def extract_product_keywords_from_title(title: str, brand: str, category: str = "", max_keywords: int = 3) -> List[str]:
    """
    네이버 쇼핑 상위 상품명에서 '브랜드 + 상품/모델' 후보를 간단 추출한다.
    추출 후보는 검색광고/쇼핑 API로 다시 검증되므로 보수적으로 적게 뽑는다.
    """
    text = clean_shopping_title(title)
    if not text:
        return []
    tokens = [t for t in re.split(r"\s+", text) if t and t not in PRODUCT_STOPWORDS]
    brand_norm = compact_keyword(brand)
    out: List[str] = []

    # 1) 브랜드 토큰 주변의 2~4그램
    for i in range(len(tokens)):
        token_norm = compact_keyword(tokens[i])
        if brand_norm and (brand_norm in token_norm or token_norm in brand_norm):
            for n in [2, 3, 4]:
                cand = " ".join(tokens[i:i+n])
                if looks_like_product_keyword(cand, brand, category):
                    out.append(cand)
            break

    # 2) 브랜드가 제목에 명확히 없으면 품목/모델 단서가 있는 2~3그램
    for i in range(len(tokens)):
        for n in [2, 3]:
            cand = " ".join(tokens[i:i+n])
            if looks_like_product_keyword(cand, brand, category):
                out.append(cand)

    # 중복 제거
    seen = set()
    uniq = []
    for cand in out:
        cn = compact_keyword(cand)
        if cn and cn not in seen and len(cand) <= 40:
            seen.add(cn)
            uniq.append(cand)
        if len(uniq) >= max_keywords:
            break
    return uniq


def discover_product_keywords(
    seed_df: pd.DataFrame,
    searchad_client: Optional[NaverSearchAdClient],
    open_client: Optional[NaverOpenApiClient] = None,
    seed_limit: int = 60,
    per_brand: int = 8,
    min_volume: int = 1000,
    max_candidates: int = 200,
    sleep_sec: float = 0.35,
    use_shopping_titles: bool = True,
    shopping_title_display: int = 30,
    exclude_risky: bool = True,
) -> pd.DataFrame:
    """
    브랜드 단위가 아니라 '브랜드 + 모델/품목' 단위 후보를 자동 발굴한다.
    주요 소스:
      1) 검색광고 keywordstool 연관키워드: 검색량이 있는 상품 키워드
      2) 네이버 쇼핑 상위 상품명: 실제 노출/판매 중인 상품명에서 모델 후보 추출
    """
    if searchad_client is None:
        print("[WARN] 검색광고 API 키가 없어 상품 단위 키워드 발굴을 건너뜁니다.", file=sys.stderr)
        return pd.DataFrame(columns=[
            "brand", "category", "keyword_series", "base_volume",
            "naver_products", "naver_overseas_products",
            "source_type", "discovery_seed", "product_group", "product_source"
        ])

    work = seed_df.copy()
    work["_base_volume_num"] = pd.to_numeric(work.get("base_volume", 0), errors="coerce").fillna(0)
    work = work.sort_values("_base_volume_num", ascending=False)
    if seed_limit and seed_limit > 0:
        work = work.head(seed_limit)

    existing = set()
    for _, r in seed_df.iterrows():
        existing.add(keyword_norm(pick_primary_keyword(r)))
        existing.add(keyword_norm(r.get("brand", "")))

    discovered_by_kw: Dict[str, Dict[str, Any]] = {}

    def add_candidate(keyword: str, brand: str, category: str, volume: int, seed: str, source: str) -> None:
        keyword = api_safe_keyword(keyword, brand, max_chars=40, max_bytes=100)
        if exclude_risky and is_too_broad_or_risky_discovery(keyword):
            return
        if not looks_like_product_keyword(keyword, brand, category):
            return
        if volume < min_volume:
            return
        kn = keyword_norm(keyword)
        if not kn or kn in existing:
            return
        if kn in discovered_by_kw:
            old = discovered_by_kw[kn]
            if volume > to_int(old.get("base_volume")):
                old["base_volume"] = volume
            return
        discovered_by_kw[kn] = {
            "brand": clean_text(brand),
            "category": product_group_from_keyword(keyword, category),
            "keyword_series": clean_text(keyword),
            "base_volume": volume,
            "naver_products": 0,
            "naver_overseas_products": 0,
            "source_type": "product_discovered",
            "discovery_seed": clean_text(seed),
            "product_group": product_group_from_keyword(keyword, category),
            "product_source": source,
        }

    for _, row in work.iterrows():
        if len(discovered_by_kw) >= max_candidates:
            break

        brand = clean_text(row.get("brand", ""))
        category = clean_text(row.get("category", ""))
        seed_keyword = api_safe_keyword(pick_primary_keyword(row), brand)
        if not seed_keyword:
            continue

        print(f"[PRODUCT DISCOVER {len(discovered_by_kw)}/{max_candidates}] {brand} -> {seed_keyword}", flush=True)

        # 1) 검색광고 연관키워드에서 상품/모델 단위 키워드 발굴
        try:
            resp = searchad_client.keyword_tool(seed_keyword)
            items = searchad_keyword_items(resp)
            ranked_items = sorted(items, key=keyword_volume_from_item, reverse=True)
            picked = 0
            for item in ranked_items:
                rel = clean_text(item.get("relKeyword", ""))
                vol = keyword_volume_from_item(item)
                if vol < min_volume:
                    continue
                if looks_like_product_keyword(rel, brand, category):
                    add_candidate(rel, brand, category, vol, seed_keyword, "searchad_related")
                    picked += 1
                if picked >= per_brand or len(discovered_by_kw) >= max_candidates:
                    break
        except Exception as e:
            print(f"[WARN] product searchad discovery failed: {brand}/{seed_keyword}: {str(e)[:120]}", file=sys.stderr)

        time.sleep(sleep_sec)

        # 2) 네이버 쇼핑 상위 상품명에서 후보 추출 후 검색광고 검색량으로 검증
        if use_shopping_titles and open_client is not None and len(discovered_by_kw) < max_candidates:
            try:
                shop = open_client.shopping_search(seed_keyword, display=shopping_title_display, sort="sim")
                title_candidates: List[str] = []
                for item in shop.get("items", []) or []:
                    title_candidates.extend(extract_product_keywords_from_title(item.get("title", ""), brand, category, max_keywords=2))

                # 후보가 너무 많아지지 않도록 순서 유지 중복제거
                seen_titles = set()
                title_candidates_unique = []
                for cand in title_candidates:
                    cn = compact_keyword(cand)
                    if cn not in seen_titles:
                        seen_titles.add(cn)
                        title_candidates_unique.append(cand)
                    if len(title_candidates_unique) >= per_brand * 2:
                        break

                for cand in title_candidates_unique:
                    if len(discovered_by_kw) >= max_candidates:
                        break
                    try:
                        vol, source = query_keyword_volume(searchad_client, cand, fallback=0, sleep_sec=sleep_sec)
                    except Exception:
                        vol, source = 0, "searchad_error"
                    if vol >= min_volume:
                        add_candidate(cand, brand, category, vol, seed_keyword, "shopping_title")
            except Exception as e:
                print(f"[WARN] shopping-title discovery failed: {brand}/{seed_keyword}: {str(e)[:120]}", file=sys.stderr)

        time.sleep(sleep_sec)

    if not discovered_by_kw:
        return pd.DataFrame(columns=[
            "brand", "category", "keyword_series", "base_volume",
            "naver_products", "naver_overseas_products",
            "source_type", "discovery_seed", "product_group", "product_source"
        ])

    out = pd.DataFrame(list(discovered_by_kw.values()))
    out = out.sort_values("base_volume", ascending=False).head(max_candidates).reset_index(drop=True)
    print(f"[INFO] 상품 단위 신규 후보 {len(out)}개 발굴 완료", flush=True)
    return out


def discover_related_keywords(
    seed_df: pd.DataFrame,
    client: Optional[NaverSearchAdClient],
    related_per_brand: int = 5,
    min_volume: int = 3000,
    seed_limit: int = 40,
    max_discovered: int = 100,
    sleep_sec: float = 0.35,
    exclude_risky: bool = True,
) -> pd.DataFrame:
    """
    기존 브랜드/키워드에서 검색광고 연관키워드를 가져와 신규 후보를 만든다.
    - 기존 엑셀에 없는 후보만 추가
    - 월간검색수 하한, 최대 개수 제한으로 GitHub Actions 실행시간 폭주 방지
    """
    if client is None:
        print("[WARN] 검색광고 API 키가 없어 신규 키워드 자동 발굴을 건너뜁니다.", file=sys.stderr)
        return pd.DataFrame(columns=[
            "brand", "category", "keyword_series", "base_volume",
            "naver_products", "naver_overseas_products",
            "source_type", "discovery_seed"
        ])

    work = seed_df.copy()
    work["_base_volume_num"] = pd.to_numeric(work.get("base_volume", 0), errors="coerce").fillna(0)
    work = work.sort_values("_base_volume_num", ascending=False)
    if seed_limit and seed_limit > 0:
        work = work.head(seed_limit)

    existing = set()
    for _, r in seed_df.iterrows():
        existing.add((keyword_norm(r.get("brand", "")), keyword_norm(pick_primary_keyword(r))))
        existing.add((keyword_norm(r.get("brand", "")), keyword_norm(r.get("brand", ""))))

    discovered_by_key: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for idx, row in work.iterrows():
        brand = clean_text(row.get("brand", ""))
        seed_keyword = api_safe_keyword(pick_primary_keyword(row), brand)
        if not seed_keyword:
            continue

        print(f"[DISCOVER {len(discovered_by_key)}/{max_discovered}] {brand} -> {seed_keyword}", flush=True)

        try:
            resp = client.keyword_tool(seed_keyword)
            items = searchad_keyword_items(resp)
        except Exception as e:
            print(f"[WARN] discovery failed: {brand}/{seed_keyword}: {str(e)[:100]}", file=sys.stderr)
            time.sleep(sleep_sec)
            continue

        candidates = []
        for item in items:
            rel = clean_text(item.get("relKeyword", ""))
            vol = keyword_volume_from_item(item)
            if not rel or vol < min_volume:
                continue
            if exclude_risky and is_too_broad_or_risky_discovery(rel):
                continue
            # API 안전 길이. 너무 긴 키워드는 신규 후보로 쓰지 않음.
            if len(rel) > 40 or len(rel.encode("utf-8")) > 100 or rel.count(" ") >= 4:
                continue

            key = (keyword_norm(brand), keyword_norm(rel))
            if key in existing:
                continue
            candidates.append((rel, vol, item))

        candidates.sort(key=lambda x: x[1], reverse=True)

        for rel, vol, item in candidates[:related_per_brand]:
            key = (keyword_norm(brand), keyword_norm(rel))
            prev = discovered_by_key.get(key)
            if prev is None or vol > prev["base_volume"]:
                discovered_by_key[key] = {
                    "brand": brand,
                    "category": clean_text(row.get("category", "")),
                    "keyword_series": rel,
                    "base_volume": int(vol),
                    "naver_products": 0,
                    "naver_overseas_products": 0,
                    "source_type": "discovered",
                    "discovery_seed": seed_keyword,
                }

            if len(discovered_by_key) >= max_discovered:
                break

        time.sleep(sleep_sec)
        if len(discovered_by_key) >= max_discovered:
            break

    if not discovered_by_key:
        return pd.DataFrame(columns=[
            "brand", "category", "keyword_series", "base_volume",
            "naver_products", "naver_overseas_products",
            "source_type", "discovery_seed"
        ])

    out = pd.DataFrame(list(discovered_by_key.values()))
    out = out.sort_values("base_volume", ascending=False).reset_index(drop=True)
    print(f"[DISCOVER DONE] 신규 후보 {len(out)}개", flush=True)
    return out


def query_keyword_volume(
    client: Optional[NaverSearchAdClient],
    keyword: str,
    fallback: int,
    retries: int = 2,
    sleep_sec: float = 0.25,
) -> Tuple[int, str]:
    if client is None:
        return fallback, "input_fallback"

    for attempt in range(retries + 1):
        try:
            resp = client.keyword_tool(keyword)
            vol, detail = parse_searchad_volume(resp, keyword)
            if vol > 0:
                rel = clean_text(detail.get("relKeyword", keyword))
                return vol, f"searchad:{rel}"
            return fallback, "searchad_empty_fallback"
        except Exception as e:
            if attempt >= retries:
                return fallback, f"searchad_error:{str(e)[:80]}"
            time.sleep(sleep_sec * (attempt + 1))

    return fallback, "input_fallback"


def query_shopping_metrics(
    client: NaverOpenApiClient,
    keyword: str,
    retries: int = 2,
    sleep_sec: float = 0.25,
    fallback_total_products: int = 0,
    fallback_overseas_products: int = 0,
) -> Dict[str, Any]:
    for attempt in range(retries + 1):
        try:
            all_data = client.shopping_search(keyword, display=10, sort="sim")
            non_cb_data = client.shopping_search(keyword, display=1, sort="sim", exclude="cbshop")
            total_products = to_int(all_data.get("total"))
            non_cb_products = to_int(non_cb_data.get("total"))
            overseas_products = max(total_products - non_cb_products, 0)

            items = all_data.get("items", []) or []
            prices = [to_int(i.get("lprice")) for i in items]
            prices = [p for p in prices if p > 0]
            avg_price = int(sum(prices) / len(prices)) if prices else 0
            min_price = min(prices) if prices else 0

            sample_titles = " | ".join(clean_text(i.get("title", "")) for i in items[:3])
            sample_malls = " | ".join(clean_text(i.get("mallName", "")) for i in items[:3])

            return {
                "total_products": total_products,
                "non_cb_products": non_cb_products,
                "overseas_products": overseas_products,
                "overseas_ratio": safe_div(overseas_products, total_products),
                "avg_top10_price": avg_price,
                "min_top10_price": min_price,
                "sample_titles": sample_titles,
                "sample_malls": sample_malls,
                "shopping_status": "ok",
            }
        except Exception as e:
            if attempt >= retries:
                return {
                    "total_products": fallback_total_products,
                    "non_cb_products": max(fallback_total_products - fallback_overseas_products, 0),
                    "overseas_products": fallback_overseas_products,
                    "overseas_ratio": safe_div(fallback_overseas_products, fallback_total_products),
                    "avg_top10_price": 0,
                    "min_top10_price": 0,
                    "sample_titles": "",
                    "sample_malls": "",
                    "shopping_status": f"fallback_input_after_error:{str(e)[:100]}",
                }
            time.sleep(sleep_sec * (attempt + 1))


# -----------------------------
# 점수 계산
# -----------------------------

RISK_KEYWORDS = {
    "식품": ["식품", "젤리", "초콜릿", "커피", "빵", "과자", "건강", "영양제", "용각산", "약", "의약", "보충제"],
    "화장품": ["화장품", "미용", "선크림", "염색약", "스프레이", "향수"],
    "전자": ["디지털", "가전", "전기", "배터리", "충전기", "무선", "카메라", "면도기"],
    "고가정품": ["명품", "주얼리", "시계", "스와로브스키", "비비안웨스트우드", "티쏘", "세이코"],
}

SEASON_KEYWORDS = {
    1: ["어그", "부츠", "장갑", "머플러", "패딩", "보온", "텀블러"],
    2: ["가방", "지갑", "새학기", "문구", "운동화"],
    3: ["러닝화", "운동화", "캠핑", "등산", "샌들"],
    4: ["러닝화", "캠핑", "등산", "양산", "우산"],
    5: ["양산", "우산", "샌들", "수영", "캠핑", "모기"],
    6: ["양산", "우산", "수영", "샌들", "모기", "쿨러"],
    7: ["양산", "우산", "수영", "샌들", "모기", "래쉬가드"],
    8: ["양산", "우산", "수영", "샌들", "모기", "래쉬가드"],
    9: ["러닝화", "등산", "캠핑", "가방", "자켓"],
    10: ["등산", "캠핑", "부츠", "자켓", "가디건"],
    11: ["어그", "부츠", "장갑", "머플러", "패딩", "선물", "텀블러"],
    12: ["선물", "지갑", "가방", "주얼리", "텀블러", "장갑", "머플러"],
}


def detect_risk(row: pd.Series) -> Tuple[int, str]:
    hay = " ".join([
        clean_text(row.get("brand", "")),
        clean_text(row.get("category", "")),
        clean_text(row.get("keyword", "")),
        clean_text(row.get("sample_titles", "")),
    ]).lower()

    reasons = []
    penalty = 0

    for label, kws in RISK_KEYWORDS.items():
        if any(k.lower() in hay for k in kws):
            reasons.append(label)
            if label == "식품":
                penalty += 25
            elif label == "화장품":
                penalty += 20
            elif label == "전자":
                penalty += 12
            elif label == "고가정품":
                penalty += 10

    # 신발/의류 사이즈 반품 리스크
    if any(k in hay for k in ["신발", "운동화", "러닝화", "의류", "원피스", "바지", "티셔츠"]):
        reasons.append("사이즈반품")
        penalty += 6

    return penalty, ",".join(sorted(set(reasons))) if reasons else ""


def season_score(row: pd.Series, today: Optional[dt.date] = None) -> int:
    today = today or dt.date.today()
    month = today.month
    hay = " ".join([
        clean_text(row.get("brand", "")),
        clean_text(row.get("category", "")),
        clean_text(row.get("keyword", "")),
        clean_text(row.get("sample_titles", "")),
    ]).lower()
    kws = SEASON_KEYWORDS.get(month, [])
    return 100 if any(k.lower() in hay for k in kws) else 45


def price_score(price: int) -> float:
    # 해외직구/구매대행 초반 추천 객단가: 3만~20만원.
    # 너무 저가는 마진 적고, 너무 고가는 관부가세/정품/반품 부담.
    if price <= 0:
        return 40
    if 30000 <= price <= 200000:
        return 100
    if 20000 <= price < 30000:
        return 75
    if 200000 < price <= 350000:
        return 70
    if price < 20000:
        return 45
    return 55


def load_latest_history(reports_dir: Path) -> Optional[pd.DataFrame]:
    files = sorted(reports_dir.glob("history_*.csv"))
    if not files:
        return None
    try:
        return pd.read_csv(files[-1])
    except Exception:
        return None


def attach_trend(df: pd.DataFrame, reports_dir: Path) -> pd.DataFrame:
    prev = load_latest_history(reports_dir)
    df["prev_search_volume"] = 0
    df["search_volume_growth_pct"] = 0.0

    if prev is None or "brand" not in prev.columns or "search_volume" not in prev.columns:
        return df

    prev_small = prev[["brand", "keyword", "search_volume"]].rename(
        columns={"search_volume": "prev_search_volume"}
    )
    merged = df.merge(prev_small, on=["brand", "keyword"], how="left", suffixes=("", "_old"))
    if "prev_search_volume_old" in merged.columns:
        merged["prev_search_volume"] = merged["prev_search_volume_old"].fillna(0)
        merged = merged.drop(columns=["prev_search_volume_old"])

    merged["prev_search_volume"] = pd.to_numeric(merged["prev_search_volume"], errors="coerce").fillna(0)
    merged["search_volume_growth_pct"] = (
        (merged["search_volume"] - merged["prev_search_volume"]) / merged["prev_search_volume"].replace(0, pd.NA) * 100
    ).fillna(0).replace([math.inf, -math.inf], 0)
    return merged


def score_dataframe(df: pd.DataFrame, reports_dir: Path) -> pd.DataFrame:
    # API 실패/무결과 시 입력 파일에 있던 네이버 상품수/해외상품수를 fallback으로 사용
    for target, fallback in [
        ("total_products", "input_total_products"),
        ("overseas_products", "input_overseas_products"),
    ]:
        if target in df.columns and fallback in df.columns:
            target_num = pd.to_numeric(df[target], errors="coerce").fillna(0)
            fallback_num = pd.to_numeric(df[fallback], errors="coerce").fillna(0)
            df[target] = target_num.where(target_num > 0, fallback_num)
    if "total_products" in df.columns and "overseas_products" in df.columns:
        df["overseas_ratio"] = df.apply(lambda r: safe_div(float(r["overseas_products"]), float(r["total_products"])), axis=1)

    df = attach_trend(df, reports_dir)

    df["demand_raw"] = pd.to_numeric(df["search_volume"], errors="coerce").fillna(0).map(lambda x: math.log1p(x))
    df["opportunity_raw"] = df.apply(
        lambda r: safe_div(float(r["search_volume"]), float(r["overseas_products"]) + 100.0),
        axis=1,
    )
    df["competition_raw"] = pd.to_numeric(df["overseas_products"], errors="coerce").fillna(0).map(lambda x: math.log1p(x))
    df["trend_raw"] = pd.to_numeric(df["search_volume_growth_pct"], errors="coerce").fillna(0).clip(-50, 200)
    df["season_score"] = df.apply(season_score, axis=1)
    df["price_score"] = df["avg_top10_price"].map(lambda x: price_score(to_int(x)))

    risk = df.apply(detect_risk, axis=1)
    df["risk_penalty"] = [x[0] for x in risk]
    df["risk_flags"] = [x[1] for x in risk]

    df["demand_score"] = pct_rank(df["demand_raw"], True)
    df["opportunity_score"] = pct_rank(df["opportunity_raw"], True)
    # 경쟁이 적을수록 좋지만, 해외직구 상품수가 0이면 시장 검증이 약하므로 하한 보정
    df["competition_score"] = pct_rank(df["competition_raw"], higher_is_better=False)
    df.loc[df["overseas_products"] <= 5, "competition_score"] *= 0.65
    df["trend_score"] = pct_rank(df["trend_raw"], True)

    df["final_score"] = (
        df["demand_score"] * 0.25
        + df["trend_score"] * 0.18
        + df["opportunity_score"] * 0.22
        + df["competition_score"] * 0.12
        + df["price_score"] * 0.13
        + df["season_score"] * 0.10
        - df["risk_penalty"]
    ).round(1)

    df["recommendation"] = pd.cut(
        df["final_score"],
        bins=[-999, 45, 60, 75, 999],
        labels=["보류", "관찰", "테스트", "우선소싱"],
    )

    cols = [
        "recommendation", "final_score",
        "source_type", "discovery_seed", "product_group", "product_source",
        "brand", "category", "keyword",
        "search_volume", "prev_search_volume", "search_volume_growth_pct",
        "total_products", "overseas_products", "overseas_ratio",
        "opportunity_raw", "avg_top10_price", "min_top10_price",
        "risk_flags", "risk_penalty",
        "volume_source", "shopping_status",
        "sample_titles", "sample_malls",
    ]
    keep = [c for c in cols if c in df.columns]
    return df.sort_values(["final_score", "search_volume"], ascending=[False, False])[keep].reset_index(drop=True)


# -----------------------------
# 리포트 저장
# -----------------------------

def save_report(df: pd.DataFrame, reports_dir: Path, top_n: int) -> Path:
    reports_dir.mkdir(parents=True, exist_ok=True)
    today = dt.date.today().strftime("%Y%m%d")
    out = reports_dir / f"sourcing_rank_{today}.xlsx"

    top_df = df.head(top_n).copy()
    source_series = df.get("source_type", pd.Series(dtype=str)).astype(str) if "source_type" in df.columns else pd.Series(dtype=str)
    discovered_count = int((source_series == "discovered").sum()) if len(source_series) else 0
    product_discovered_count = int((source_series == "product_discovered").sum()) if len(source_series) else 0
    seed_count = int((source_series == "seed").sum()) if len(source_series) else len(df)
    summary = pd.DataFrame({
        "항목": [
            "생성일",
            "전체 후보 수",
            "기존 후보 수",
            "신규 연관키워드 후보 수",
            "상품 단위 발굴 후보 수",
            f"TOP {top_n} 평균 점수",
            "우선소싱 수",
            "테스트 수",
            "보류 수",
            "주의",
        ],
        "값": [
            dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            len(df),
            seed_count,
            discovered_count,
            product_discovered_count,
            round(float(top_df["final_score"].mean()), 1) if len(top_df) else 0,
            int((df["recommendation"].astype(str) == "우선소싱").sum()),
            int((df["recommendation"].astype(str) == "테스트").sum()),
            int((df["recommendation"].astype(str) == "보류").sum()),
            "점수는 소싱 후보 선별용입니다. 식품/화장품/KC/상표권/정품 증빙은 등록 전 별도 확인하세요.",
        ],
    })

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="요약")
        top_df.to_excel(writer, index=False, sheet_name=f"TOP_{top_n}")
        if "source_type" in df.columns:
            new_df = df[df["source_type"].astype(str) == "discovered"].copy()
            new_df.head(top_n).to_excel(writer, index=False, sheet_name="신규발굴_TOP")
            product_df = df[df["source_type"].astype(str) == "product_discovered"].copy()
            product_df.head(top_n).to_excel(writer, index=False, sheet_name="상품키워드_TOP")
        df.to_excel(writer, index=False, sheet_name="전체랭킹")

        workbook = writer.book
        fmt_header = workbook.add_format({"bold": True, "bg_color": "#E8F4F8", "border": 1})
        fmt_money = workbook.add_format({"num_format": "#,##0"})
        fmt_pct = workbook.add_format({"num_format": "0.0%"})
        fmt_score = workbook.add_format({"num_format": "0.0"})
        fmt_warn = workbook.add_format({"bg_color": "#FFF2CC"})
        fmt_good = workbook.add_format({"bg_color": "#D9EAD3"})
        fmt_bad = workbook.add_format({"bg_color": "#F4CCCC"})

        report_sheets = ["요약", f"TOP_{top_n}", "전체랭킹"]
        insert_at = 2
        if "신규발굴_TOP" in writer.sheets:
            report_sheets.insert(insert_at, "신규발굴_TOP")
            insert_at += 1
        if "상품키워드_TOP" in writer.sheets:
            report_sheets.insert(insert_at, "상품키워드_TOP")
        for sheet_name in report_sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(1, len(df)), max(1, len(df.columns) - 1))
            ws.set_row(0, None, fmt_header)
            ws.set_column(0, 0, 12)
            ws.set_column(1, 1, 11, fmt_score)
            ws.set_column(2, 4, 18)
            ws.set_column(5, 7, 16, fmt_money)
            ws.set_column(8, 10, 15, fmt_money)
            ws.set_column(11, 12, 14, fmt_money)
            ws.set_column(13, 14, 14)
            ws.set_column(15, 18, 30)

        ws = writer.sheets[f"TOP_{top_n}"]
        end_row = len(top_df)
        if end_row > 0:
            rec_col = 0
            ws.conditional_format(1, rec_col, end_row, rec_col, {
                "type": "text", "criteria": "containing", "value": "우선소싱", "format": fmt_good
            })
            ws.conditional_format(1, rec_col, end_row, rec_col, {
                "type": "text", "criteria": "containing", "value": "보류", "format": fmt_bad
            })
            risk_col = top_df.columns.get_loc("risk_flags") if "risk_flags" in top_df.columns else None
            if risk_col is not None:
                ws.conditional_format(1, risk_col, end_row, risk_col, {
                    "type": "no_blanks", "format": fmt_warn
                })

    hist = reports_dir / f"history_{today}.csv"
    df.to_csv(hist, index=False, encoding="utf-8-sig")
    if "source_type" in df.columns:
        discovered = df[df["source_type"].astype(str) == "discovered"].copy()
        if len(discovered):
            discovered.to_csv(reports_dir / f"discovered_keywords_{today}.csv", index=False, encoding="utf-8-sig")
        product_discovered = df[df["source_type"].astype(str) == "product_discovered"].copy()
        if len(product_discovered):
            product_discovered.to_csv(reports_dir / f"product_keywords_{today}.csv", index=False, encoding="utf-8-sig")
    return out


# -----------------------------
# 메인
# -----------------------------

def build_clients(args: argparse.Namespace) -> Tuple[NaverOpenApiClient, Optional[NaverSearchAdClient]]:
    open_id = os.getenv("NAVER_CLIENT_ID", "").strip()
    open_secret = os.getenv("NAVER_CLIENT_SECRET", "").strip()
    if not open_id or not open_secret:
        raise RuntimeError("NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 환경변수가 필요합니다.")

    open_client = NaverOpenApiClient(open_id, open_secret)

    if args.no_searchad:
        return open_client, None

    ad_key = os.getenv("NAVER_SEARCHAD_API_KEY", "").strip()
    ad_secret = os.getenv("NAVER_SEARCHAD_SECRET_KEY", "").strip()
    ad_customer = os.getenv("NAVER_SEARCHAD_CUSTOMER_ID", "").strip()

    if not (ad_key and ad_secret and ad_customer):
        print("[WARN] 검색광고 API 키가 없어 입력 파일의 검색량을 fallback으로 사용합니다.", file=sys.stderr)
        return open_client, None

    return open_client, NaverSearchAdClient(ad_key, ad_secret, ad_customer)


def run(args: argparse.Namespace) -> Path:
    load_dotenv(args.env)

    input_path = Path(args.input)
    reports_dir = Path(args.reports_dir)
    df = load_input(input_path)

    if args.limit:
        df = df.head(args.limit).copy()

    open_client, searchad_client = build_clients(args)

    if args.discover_related:
        discovered_df = discover_related_keywords(
            df,
            searchad_client,
            related_per_brand=args.related_per_brand,
            min_volume=args.min_discovered_volume,
            seed_limit=args.discovery_seed_limit,
            max_discovered=args.max_discovered,
            sleep_sec=args.sleep,
            exclude_risky=not args.include_risky_discovery,
        )
        if len(discovered_df):
            df = pd.concat([df, discovered_df], ignore_index=True)
            # 같은 브랜드+대표키워드 중복 제거: 기존 seed 우선
            df["_priority"] = df["source_type"].map({"seed": 0, "discovered": 1}).fillna(9)
            df["_kw_norm"] = df.apply(lambda r: keyword_norm(pick_primary_keyword(r)), axis=1)
            df["_brand_norm"] = df["brand"].map(keyword_norm)
            df = df.sort_values(["_priority"]).drop_duplicates(subset=["_brand_norm", "_kw_norm"])
            df = df.drop(columns=["_priority", "_kw_norm", "_brand_norm"]).reset_index(drop=True)

    if args.discover_products:
        product_df = discover_product_keywords(
            df,
            searchad_client,
            open_client=open_client,
            seed_limit=args.product_seed_limit,
            per_brand=args.product_keywords_per_brand,
            min_volume=args.min_product_volume,
            max_candidates=args.max_product_candidates,
            sleep_sec=args.sleep,
            use_shopping_titles=not args.no_shopping_title_discovery,
            shopping_title_display=args.shopping_title_display,
            exclude_risky=not args.include_risky_discovery,
        )
        if len(product_df):
            df = pd.concat([df, product_df], ignore_index=True)

    # 전체 중복 제거: 기존 seed를 우선 보존하고, 상품 단위 발굴 후보도 같은 키워드는 1개만 유지
    if "source_type" in df.columns:
        df["_priority"] = df["source_type"].map({"seed": 0, "discovered": 1, "product_discovered": 2}).fillna(9)
        df["_kw_norm"] = df.apply(lambda r: keyword_norm(pick_primary_keyword(r)), axis=1)
        df["_brand_norm"] = df["brand"].map(keyword_norm)
        df = df.sort_values(["_priority"]).drop_duplicates(subset=["_brand_norm", "_kw_norm"])
        df = df.drop(columns=["_priority", "_kw_norm", "_brand_norm"]).reset_index(drop=True)

    if args.products_only:
        df = df[df["source_type"].astype(str) == "product_discovered"].copy().reset_index(drop=True)
        if len(df) == 0:
            raise RuntimeError("products-only 모드인데 상품 단위 후보가 없습니다. 검색광고 API 키와 discovery 옵션을 확인하세요.")

    rows: List[Dict[str, Any]] = []
    for idx, row in df.iterrows():
        brand = clean_text(row["brand"])
        keyword = pick_primary_keyword(row)
        fallback_volume = to_int(row.get("base_volume"))

        if not keyword:
            continue

        api_keyword = api_safe_keyword(keyword, brand)
        input_total_products = to_int(row.get("naver_products"))
        input_overseas_products = to_int(row.get("naver_overseas_products"))

        print(f"[{idx + 1}/{len(df)}] {brand} / {api_keyword}", flush=True)

        if clean_text(row.get("source_type", "")) == "discovered" and fallback_volume > 0:
            # discovery 단계에서 이미 검색광고 월간검색수를 확보했으므로 중복 호출 방지
            search_volume, volume_source = fallback_volume, "searchad_discovery"
        else:
            search_volume, volume_source = query_keyword_volume(
                searchad_client, api_keyword, fallback=fallback_volume, sleep_sec=args.sleep
            )

        shopping = query_shopping_metrics(
            open_client,
            api_keyword,
            sleep_sec=args.sleep,
            fallback_total_products=input_total_products,
            fallback_overseas_products=input_overseas_products,
        )

        rows.append({
            "source_type": clean_text(row.get("source_type", "seed")) or "seed",
            "discovery_seed": clean_text(row.get("discovery_seed", "")),
            "product_group": clean_text(row.get("product_group", "")),
            "product_source": clean_text(row.get("product_source", "")),
            "brand": brand,
            "category": clean_text(row.get("category", "")),
            "keyword": api_keyword,
            "search_volume": search_volume,
            "input_search_volume": fallback_volume,
            "input_total_products": input_total_products,
            "input_overseas_products": input_overseas_products,
            "volume_source": volume_source,
            **shopping,
        })

        time.sleep(args.sleep)

    result = pd.DataFrame(rows)
    ranked = score_dataframe(result, reports_dir)
    out = save_report(ranked, reports_dir, args.top)
    print(f"\n[DONE] {out}")
    return out


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="네이버 해외직구/구매대행 소싱 후보 자동 랭킹")
    p.add_argument("--input", required=True, help="브랜드/키워드 후보 엑셀 또는 CSV")
    p.add_argument("--env", default=".env", help=".env 파일 경로")
    p.add_argument("--reports-dir", default="reports", help="리포트 저장 폴더")
    p.add_argument("--top", type=int, default=80, help="TOP 시트에 넣을 개수")
    p.add_argument("--limit", type=int, default=0, help="테스트용 조회 개수 제한")
    p.add_argument("--sleep", type=float, default=0.35, help="API 호출 간격 초")
    p.add_argument("--no-searchad", action="store_true", help="검색광고 API 호출 없이 입력 파일 검색량 사용")
    p.add_argument("--discover-related", action="store_true", help="검색광고 연관키워드로 신규 후보 자동 발굴")
    p.add_argument("--related-per-brand", type=int, default=5, help="브랜드/시드당 추가할 연관키워드 개수")
    p.add_argument("--min-discovered-volume", type=int, default=3000, help="신규 후보 최소 월간검색수")
    p.add_argument("--discovery-seed-limit", type=int, default=40, help="신규 발굴에 사용할 상위 시드 개수")
    p.add_argument("--max-discovered", type=int, default=100, help="최대 신규 발굴 후보 수")
    p.add_argument("--include-risky-discovery", action="store_true", help="규제/리스크 키워드도 신규 후보에 포함")
    p.add_argument("--discover-products", action="store_true", help="브랜드가 아니라 제품/모델 단위 키워드를 자동 발굴")
    p.add_argument("--products-only", action="store_true", help="최종 랭킹에서 브랜드 seed를 제외하고 상품 단위 발굴 후보만 출력")
    p.add_argument("--product-seed-limit", type=int, default=60, help="상품 단위 발굴에 사용할 상위 브랜드/시드 개수")
    p.add_argument("--product-keywords-per-brand", type=int, default=8, help="브랜드당 발굴할 상품 키워드 최대 개수")
    p.add_argument("--min-product-volume", type=int, default=1000, help="상품 단위 후보 최소 월간검색수")
    p.add_argument("--max-product-candidates", type=int, default=200, help="최대 상품 단위 신규 후보 수")
    p.add_argument("--no-shopping-title-discovery", action="store_true", help="네이버 쇼핑 상위 상품명 기반 상품 후보 추출을 끔")
    p.add_argument("--shopping-title-display", type=int, default=30, help="상품명 후보 추출에 사용할 쇼핑 상위 노출 개수")
    return p.parse_args(argv)


if __name__ == "__main__":
    run(parse_args())
