#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
네이버 해외직구/구매대행 소싱 후보 자동 랭킹
- 입력: 브랜드&키워드 엑셀/CSV
- 조회: 네이버 검색광고 키워드도구(월간검색수), 네이버 쇼핑검색 API(전체/해외상품수, 국내가 샘플)
- 출력: reports/sourcing_rank_YYYYMMDD.xlsx, history snapshot CSV

필수 환경변수:
  NAVER_CLIENT_ID
  NAVER_CLIENT_SECRET

선택 환경변수(월간 검색량 자동 갱신용):
  NAVER_SEARCHAD_API_KEY
  NAVER_SEARCHAD_SECRET_KEY
  NAVER_SEARCHAD_CUSTOMER_ID

사용 예:
  python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 80
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
    for cand in COLUMN_ALIASES[logical]:
        if cand in df.columns:
            return cand
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
    df["keyword_series"] = df["keyword_series"].map(clean_text)
    df = df[df["brand"] != ""].drop_duplicates(subset=["brand", "keyword_series"])
    return df.reset_index(drop=True)


def pick_primary_keyword(row: pd.Series) -> str:
    raw = clean_text(row.get("keyword_series", "")) or clean_text(row.get("brand", ""))
    # 파일 안에서 줄바꿈/쉼표로 시리즈 키워드가 여러 개 들어온 경우 첫 번째를 대표로 사용
    parts = re.split(r"[\n\r,;/|]+", raw)
    parts = [p.strip() for p in parts if p.strip()]
    return parts[0] if parts else clean_text(row.get("brand", ""))


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
                    "total_products": 0,
                    "non_cb_products": 0,
                    "overseas_products": 0,
                    "overseas_ratio": 0,
                    "avg_top10_price": 0,
                    "min_top10_price": 0,
                    "sample_titles": "",
                    "sample_malls": "",
                    "shopping_status": f"error:{str(e)[:100]}",
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
    summary = pd.DataFrame({
        "항목": [
            "생성일",
            "전체 후보 수",
            f"TOP {top_n} 평균 점수",
            "우선소싱 수",
            "테스트 수",
            "보류 수",
            "주의",
        ],
        "값": [
            dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            len(df),
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
        df.to_excel(writer, index=False, sheet_name="전체랭킹")

        workbook = writer.book
        fmt_header = workbook.add_format({"bold": True, "bg_color": "#E8F4F8", "border": 1})
        fmt_money = workbook.add_format({"num_format": "#,##0"})
        fmt_pct = workbook.add_format({"num_format": "0.0%"})
        fmt_score = workbook.add_format({"num_format": "0.0"})
        fmt_warn = workbook.add_format({"bg_color": "#FFF2CC"})
        fmt_good = workbook.add_format({"bg_color": "#D9EAD3"})
        fmt_bad = workbook.add_format({"bg_color": "#F4CCCC"})

        for sheet_name in ["요약", f"TOP_{top_n}", "전체랭킹"]:
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

    rows: List[Dict[str, Any]] = []
    for idx, row in df.iterrows():
        brand = clean_text(row["brand"])
        keyword = pick_primary_keyword(row)
        fallback_volume = to_int(row.get("base_volume"))

        if not keyword:
            continue

        print(f"[{idx + 1}/{len(df)}] {brand} / {keyword}", flush=True)

        search_volume, volume_source = query_keyword_volume(
            searchad_client, keyword, fallback=fallback_volume, sleep_sec=args.sleep
        )

        shopping = query_shopping_metrics(open_client, keyword, sleep_sec=args.sleep)

        rows.append({
            "brand": brand,
            "category": clean_text(row.get("category", "")),
            "keyword": keyword,
            "search_volume": search_volume,
            "input_search_volume": fallback_volume,
            "input_total_products": to_int(row.get("naver_products")),
            "input_overseas_products": to_int(row.get("naver_overseas_products")),
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
    return p.parse_args(argv)


if __name__ == "__main__":
    run(parse_args())
