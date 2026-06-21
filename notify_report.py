#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Actions 실행 후 네이버 소싱 랭킹 결과를 요약하고 발송합니다.

지원 채널:
- GitHub Step Summary: 항상 생성
- Telegram: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID가 있으면 발송
- Slack/Discord/Make/Zapier Webhook: REPORT_WEBHOOK_URL이 있으면 POST
- Email SMTP: SMTP_HOST, SMTP_USER, SMTP_PASSWORD, REPORT_TO_EMAIL이 있으면 발송
- Gemini 요약: GEMINI_API_KEY가 있으면 TOP 데이터를 더 자연스러운 한국어 브리핑으로 요약

주의:
- API 키/토큰은 절대 코드에 직접 적지 말고 GitHub Secrets에 저장하세요.
"""

from __future__ import annotations

import argparse
import os
import glob
import json
import smtplib
import textwrap
from email.message import EmailMessage
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from dotenv import load_dotenv


def find_latest_report(reports_dir: str = "reports") -> Path:
    files = sorted(glob.glob(f"{reports_dir}/sourcing_rank_*.xlsx"))
    if not files:
        files = sorted(glob.glob(f"{reports_dir}/*.xlsx"))
    if not files:
        raise FileNotFoundError(f"{reports_dir} 폴더에서 엑셀 리포트를 찾지 못했습니다.")
    return Path(files[-1])


def pick_sheet(path: Path) -> str:
    xls = pd.ExcelFile(path)
    for name in ["TOP_80", "TOP", "전체랭킹", "요약"]:
        if name in xls.sheet_names:
            return name
    return xls.sheet_names[0]


def read_top_rows(path: Path, top_n: int = 15) -> pd.DataFrame:
    sheet = pick_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet)
    # 보기 좋은 컬럼만 남기기. 컬럼명이 다소 바뀌어도 최대한 대응.
    preferred = [
        "rank", "순위", "recommendation", "추천등급", "sourcing_score", "최종점수", "final_score",
        "source_type", "discovery_seed",
        "brand", "브랜드명", "primary_keyword", "대표키워드", "keyword",
        "category", "대표카테고리", "search_volume", "검색량",
        "volume_growth_pct", "상승률", "search_volume_growth_pct",
        "naver_products", "네이버상품수", "total_products",
        "naver_overseas_products", "네이버해외상품수", "overseas_products",
        "avg_price", "평균가", "avg_top10_price", "risk_flags", "리스크"
    ]
    cols = [c for c in preferred if c in df.columns]
    if cols:
        df = df[cols].copy()
    return df.head(top_n)


def df_to_compact_text(df: pd.DataFrame, top_n: int = 10) -> str:
    lines = []
    for idx, row in df.head(top_n).iterrows():
        def val(*names, default=""):
            for n in names:
                if n in row and pd.notna(row[n]):
                    return row[n]
            return default

        rank = val("rank", "순위", default=idx + 1)
        brand = val("brand", "브랜드명")
        keyword = val("primary_keyword", "대표키워드", "keyword", default=brand)
        score = val("sourcing_score", "최종점수", "final_score", default="")
        volume = val("search_volume", "검색량", default="")
        overseas = val("naver_overseas_products", "네이버해외상품수", "overseas_products", default="")
        risk = val("risk_flags", "리스크", default="")
        source_type = val("source_type", default="")
        source_txt = "신규발굴" if str(source_type) == "discovered" else "기존"
        score_txt = f" / 점수 {float(score):.1f}" if isinstance(score, (int, float)) else (f" / 점수 {score}" if score != "" else "")
        volume_txt = f" / 검색량 {int(volume):,}" if isinstance(volume, (int, float)) else (f" / 검색량 {volume}" if volume != "" else "")
        overseas_txt = f" / 해외상품 {int(overseas):,}" if isinstance(overseas, (int, float)) else (f" / 해외상품 {overseas}" if overseas != "" else "")
        risk_txt = f" / 주의 {risk}" if risk else ""
        lines.append(f"{rank}. [{source_txt}] {brand} - {keyword}{score_txt}{volume_txt}{overseas_txt}{risk_txt}")
    return "\n".join(lines)


def gemini_summary(raw_text: str) -> Optional[str]:
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None

    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        prompt = f"""
너는 한국 스마트스토어 해외직구/사입 MD다.
아래 자동 랭킹 TOP 데이터를 보고, 오늘 바로 확인할 소싱 액션 브리핑을 한국어로 작성해줘.

규칙:
- 과장 금지
- 식품/화장품/의약외품/전자제품/KC/상표권 리스크가 보이면 주의 문구 포함
- TOP 5만 짧게 추천
- 각 항목마다 '추천 아이템 방향'과 '확인할 것'을 적어줘
- 전체 900자 이내

데이터:
{raw_text}
"""
        response = client.models.generate_content(
            model=os.getenv("GEMINI_MODEL", "gemini-3.5-flash"),
            contents=prompt,
        )
        return getattr(response, "text", None)
    except Exception as e:
        return f"[Gemini 요약 실패: {e}]\n\n{raw_text}"


def build_message(report_path: Path, top_n: int = 15) -> str:
    df = read_top_rows(report_path, top_n=top_n)
    raw = df_to_compact_text(df, top_n=min(top_n, 15))
    ai = gemini_summary(raw)
    if ai:
        body = ai.strip() + "\n\n---\n원본 TOP 데이터\n" + raw
    else:
        body = "이번 실행의 소싱 후보 TOP 데이터입니다.\n\n" + raw

    return f"""📦 네이버 해외직구 소싱 자동 리포트

리포트 파일: {report_path.name}

{body}

GitHub Actions의 Artifacts에서 엑셀 파일을 내려받을 수 있습니다.
"""


def write_github_summary(message: str, report_path: Path) -> None:
    summary_path = os.getenv("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return
    with open(summary_path, "a", encoding="utf-8") as f:
        f.write("## 네이버 해외직구 소싱 자동 리포트\n\n")
        f.write(message.replace("\n", "\n\n"))
        f.write("\n\n")
        f.write(f"- Excel report: `{report_path}`\n")


def send_telegram(message: str) -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        return
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    # 텔레그램 메시지 길이 제한을 고려해 잘라서 발송
    chunks = [message[i:i+3500] for i in range(0, len(message), 3500)]
    for chunk in chunks:
        r = requests.post(url, json={"chat_id": chat_id, "text": chunk}, timeout=15)
        r.raise_for_status()


def send_webhook(message: str) -> None:
    url = os.getenv("REPORT_WEBHOOK_URL", "").strip()
    if not url:
        return
    payload = {"text": message, "content": message}
    r = requests.post(url, json=payload, timeout=15)
    r.raise_for_status()


def send_email(message: str, report_path: Path) -> None:
    host = os.getenv("SMTP_HOST", "").strip()
    user = os.getenv("SMTP_USER", "").strip()
    password = os.getenv("SMTP_PASSWORD", "").strip()
    to_email = os.getenv("REPORT_TO_EMAIL", "").strip()
    if not host or not user or not password or not to_email:
        return

    port = int(os.getenv("SMTP_PORT", "587"))
    from_email = os.getenv("REPORT_FROM_EMAIL", user)
    subject = os.getenv("REPORT_EMAIL_SUBJECT", "네이버 해외직구 소싱 자동 리포트")

    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(message)

    # 엑셀 첨부
    data = report_path.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=report_path.name,
    )

    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser()
    parser.add_argument("--reports-dir", default="reports")
    parser.add_argument("--top", type=int, default=15)
    args = parser.parse_args()

    report_path = find_latest_report(args.reports_dir)
    message = build_message(report_path, top_n=args.top)
    write_github_summary(message, report_path)

    errors = []
    for fn in [send_telegram, send_webhook]:
        try:
            fn(message)
        except Exception as e:
            errors.append(f"{fn.__name__}: {e}")

    try:
        send_email(message, report_path)
    except Exception as e:
        errors.append(f"send_email: {e}")

    print(message)
    if errors:
        print("\n알림 발송 중 일부 실패:")
        for e in errors:
            print("-", e)


if __name__ == "__main__":
    main()
