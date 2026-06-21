#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Actions 실행 후 네이버 소싱 랭킹 결과를 요약하고 발송합니다.

업그레이드 기능:
- Telegram 텍스트 요약 발송
- Telegram용 TOP 요약 이미지 자동 생성 후 sendPhoto 발송
- Streamlit 대시보드 링크 포함
- GitHub Step Summary, Webhook, Email 지원

필수:
- TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID를 GitHub Secrets에 저장
- DASHBOARD_URL 또는 STREAMLIT_APP_URL을 GitHub Variables에 저장하면 메시지에 링크가 포함됨

주의:
- API 키/토큰은 절대 코드에 직접 적지 말고 GitHub Secrets에 저장하세요.
"""

from __future__ import annotations

import argparse
import glob
import os
import smtplib
import textwrap
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import requests
from dotenv import load_dotenv


# -----------------------------
# 리포트 로딩
# -----------------------------

def find_latest_report(reports_dir: str = "reports") -> Path:
    files = sorted(glob.glob(f"{reports_dir}/sourcing_rank_*.xlsx"))
    if not files:
        files = sorted(glob.glob(f"{reports_dir}/*.xlsx"))
    if not files:
        raise FileNotFoundError(f"{reports_dir} 폴더에서 엑셀 리포트를 찾지 못했습니다.")
    return Path(files[-1])


def pick_sheet(path: Path) -> str:
    xls = pd.ExcelFile(path)
    # 상품 단위 발굴 버전 우선
    for name in ["상품키워드_TOP", "신규발굴_TOP", "TOP_100", "TOP_80", "TOP", "전체랭킹", "요약"]:
        if name in xls.sheet_names:
            return name
    return xls.sheet_names[0]


def read_top_rows(path: Path, top_n: int = 15) -> pd.DataFrame:
    sheet = pick_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet)

    preferred = [
        "rank", "순위",
        "recommendation", "추천등급",
        "sourcing_score", "최종점수", "final_score", "score", "소싱점수",
        "source_type", "discovery_seed", "product_group", "product_source",
        "brand", "브랜드명",
        "primary_keyword", "대표키워드", "keyword", "키워드",
        "category", "대표카테고리",
        "search_volume", "검색량",
        "volume_growth_pct", "상승률", "search_volume_growth_pct",
        "naver_products", "네이버상품수", "total_products",
        "naver_overseas_products", "네이버해외상품수", "overseas_products",
        "avg_price", "평균가", "avg_top10_price", "국내평균가",
        "min_price", "최저가", "min_top10_price", "국내최저가",
        "risk_flags", "리스크"
    ]
    cols = [c for c in preferred if c in df.columns]
    if cols:
        df = df[cols].copy()
    return df.head(top_n)


# -----------------------------
# 값 추출/포맷
# -----------------------------

def row_val(row: pd.Series, *names: str, default: Any = "") -> Any:
    for n in names:
        if n in row and pd.notna(row[n]):
            return row[n]
    return default


def to_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or pd.isna(value):
            return default
    except Exception:
        pass
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if value in {"", "nan", "None", "-"}:
            return default
        if value.startswith("<"):
            return default
    try:
        return float(value)
    except Exception:
        return default


def fmt_int(value: Any, suffix: str = "") -> str:
    num = to_float(value, 0)
    if num <= 0:
        return "-"
    return f"{int(num):,}{suffix}"


def fmt_score(value: Any) -> str:
    num = to_float(value, -1)
    if num < 0:
        return "-"
    return f"{num:.1f}"


def truncate(text: Any, max_len: int) -> str:
    s = str(text or "").strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + "…"


def dashboard_url() -> str:
    return (
        os.getenv("DASHBOARD_URL", "").strip()
        or os.getenv("STREAMLIT_APP_URL", "").strip()
        or os.getenv("REPORT_DASHBOARD_URL", "").strip()
    )


def df_to_compact_text(df: pd.DataFrame, top_n: int = 10) -> str:
    lines = []
    for idx, row in df.head(top_n).iterrows():
        rank = row_val(row, "rank", "순위", default=idx + 1)
        brand = row_val(row, "brand", "브랜드명")
        keyword = row_val(row, "primary_keyword", "대표키워드", "keyword", "키워드", default=brand)
        score = row_val(row, "sourcing_score", "최종점수", "final_score", "score", "소싱점수", default="")
        volume = row_val(row, "search_volume", "검색량", default="")
        avg_price = row_val(row, "avg_price", "평균가", "avg_top10_price", "국내평균가", default="")
        overseas = row_val(row, "naver_overseas_products", "네이버해외상품수", "overseas_products", default="")
        risk = row_val(row, "risk_flags", "리스크", default="")
        source_type = row_val(row, "source_type", default="")

        source_txt = "상품발굴" if str(source_type) == "product_discovered" else ("신규발굴" if str(source_type) == "discovered" else "기존")
        score_txt = f" / 점수 {fmt_score(score)}" if score != "" else ""
        volume_txt = f" / 검색량 {fmt_int(volume)}" if volume != "" else ""
        price_txt = f" / 평균가 {fmt_int(avg_price, '원')}" if avg_price != "" else ""
        overseas_txt = f" / 해외상품 {fmt_int(overseas)}" if overseas != "" else ""
        risk_txt = f" / 주의 {risk}" if risk else ""

        if brand and keyword and str(brand).strip() != str(keyword).strip():
            title = f"{brand} - {keyword}"
        else:
            title = str(keyword or brand or "").strip()

        lines.append(f"{rank}. [{source_txt}] {title}{score_txt}{volume_txt}{price_txt}{overseas_txt}{risk_txt}")
    return "\n".join(lines)


# -----------------------------
# Gemini 요약
# -----------------------------

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

    dash = dashboard_url()
    dash_line = f"\n\n🔗 대시보드 보기:\n{dash}" if dash else ""

    return f"""📦 네이버 해외직구 소싱 자동 리포트

리포트 파일: {report_path.name}

{body}{dash_line}

GitHub Actions의 Artifacts에서 엑셀 파일을 내려받을 수 있습니다.
"""


# -----------------------------
# 텔레그램용 요약 이미지 생성
# -----------------------------

def find_font(bold: bool = False, size: int = 30):
    """GitHub Actions Ubuntu에서 fonts-noto-cjk 설치 시 한글이 정상 표시됩니다."""
    from PIL import ImageFont

    candidates = []
    if bold:
        candidates += [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJKkr-Bold.otf",
            "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
        ]
    candidates += [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKkr-Regular.otf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]

    for path in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                pass
    return ImageFont.load_default()


def create_summary_image(report_path: Path, top_n: int = 10, output_dir: str = "reports") -> Optional[Path]:
    """엑셀 TOP 데이터를 읽어 모바일 텔레그램에서 보기 좋은 이미지로 저장."""
    try:
        from PIL import Image, ImageDraw
    except Exception as e:
        print(f"[이미지 생성 건너뜀] pillow import 실패: {e}")
        return None

    try:
        df = read_top_rows(report_path, top_n=top_n)
    except Exception as e:
        print(f"[이미지 생성 건너뜀] 리포트 읽기 실패: {e}")
        return None

    if df.empty:
        return None

    width = 1200
    top_rows = min(top_n, len(df))
    header_h = 185
    row_h = 118
    footer_h = 110 if dashboard_url() else 72
    height = header_h + row_h * top_rows + footer_h

    bg = "#F6F8FB"
    card = "#FFFFFF"
    text_main = "#111827"
    text_sub = "#4B5563"
    line = "#E5E7EB"
    accent = "#009698"
    warn = "#B45309"
    danger = "#B91C1C"

    img = Image.new("RGB", (width, height), bg)
    draw = ImageDraw.Draw(img)

    f_title = find_font(True, 44)
    f_sub = find_font(False, 24)
    f_head = find_font(True, 26)
    f_row_title = find_font(True, 30)
    f_row = find_font(False, 23)
    f_small = find_font(False, 20)
    f_rank = find_font(True, 34)

    # Header
    draw.rounded_rectangle((40, 32, width - 40, 150), radius=26, fill=card)
    draw.rectangle((40, 32, 58, 150), fill=accent)
    draw.text((80, 52), "이번 주 해외직구 소싱 TOP", font=f_title, fill=text_main)
    draw.text((82, 108), f"{report_path.name} · 상품/모델 키워드 중심", font=f_sub, fill=text_sub)

    # Column hints
    y = header_h - 32
    draw.text((86, y), "상품키워드", font=f_head, fill=text_sub)
    draw.text((625, y), "점수", font=f_head, fill=text_sub)
    draw.text((735, y), "검색량", font=f_head, fill=text_sub)
    draw.text((895, y), "평균가", font=f_head, fill=text_sub)
    draw.text((1040, y), "해외상품", font=f_head, fill=text_sub)

    # Rows
    y0 = header_h
    for i, (_, row) in enumerate(df.head(top_n).iterrows(), start=1):
        y1 = y0 + (i - 1) * row_h
        x1, x2 = 40, width - 40
        draw.rounded_rectangle((x1, y1, x2, y1 + row_h - 14), radius=20, fill=card)
        draw.line((70, y1 + row_h - 14, width - 70, y1 + row_h - 14), fill=line, width=1)

        brand = row_val(row, "brand", "브랜드명", default="")
        keyword = row_val(row, "primary_keyword", "대표키워드", "keyword", "키워드", default=brand)
        category = row_val(row, "category", "대표카테고리", default="")
        risk = row_val(row, "risk_flags", "리스크", default="")
        score = row_val(row, "sourcing_score", "최종점수", "final_score", "score", "소싱점수", default="")
        volume = row_val(row, "search_volume", "검색량", default="")
        avg_price = row_val(row, "avg_price", "평균가", "avg_top10_price", "국내평균가", default="")
        overseas = row_val(row, "naver_overseas_products", "네이버해외상품수", "overseas_products", default="")

        title = str(keyword or brand or "").strip()
        if brand and keyword and str(brand).strip() != str(keyword).strip() and str(brand).strip() not in str(keyword):
            title = f"{brand} {keyword}"

        rank_color = accent if i <= 3 else "#6B7280"
        draw.text((74, y1 + 28), f"{i}", font=f_rank, fill=rank_color)

        title_x = 128
        draw.text((title_x, y1 + 24), truncate(title, 26), font=f_row_title, fill=text_main)
        sub_parts = [str(category).strip()] if category else []
        if risk:
            sub_parts.append(f"주의: {truncate(risk, 18)}")
        sub = " · ".join(sub_parts) if sub_parts else "상품 키워드 후보"
        sub_color = danger if risk else text_sub
        draw.text((title_x, y1 + 66), truncate(sub, 36), font=f_small, fill=sub_color)

        draw.text((625, y1 + 38), fmt_score(score), font=f_row_title, fill=accent)
        draw.text((735, y1 + 42), fmt_int(volume), font=f_row, fill=text_main)
        draw.text((895, y1 + 42), fmt_int(avg_price, "원"), font=f_row, fill=text_main)
        draw.text((1040, y1 + 42), fmt_int(overseas), font=f_row, fill=text_main)

    # Footer
    fy = header_h + row_h * top_rows + 18
    dash = dashboard_url()
    if dash:
        draw.rounded_rectangle((40, fy, width - 40, fy + 76), radius=18, fill="#EFFFFB")
        draw.text((70, fy + 16), "자세한 필터링/전체 후보는 Streamlit 대시보드에서 확인하세요.", font=f_sub, fill=text_main)
        draw.text((70, fy + 47), truncate(dash, 85), font=f_small, fill=accent)
    else:
        draw.text((60, fy + 12), "대시보드 링크를 넣으려면 GitHub Variables에 DASHBOARD_URL을 추가하세요.", font=f_small, fill=text_sub)

    Path(output_dir).mkdir(parents=True, exist_ok=True)
    out_path = Path(output_dir) / f"telegram_summary_{report_path.stem.replace('sourcing_rank_', '')}.png"
    img.save(out_path, quality=95)
    return out_path


# -----------------------------
# 발송 채널
# -----------------------------

def write_github_summary(message: str, report_path: Path, image_path: Optional[Path] = None) -> None:
    summary_path = os.getenv("GITHUB_STEP_SUMMARY")
    if not summary_path:
        return
    with open(summary_path, "a", encoding="utf-8") as f:
        f.write("## 네이버 해외직구 소싱 자동 리포트\n\n")
        if image_path:
            f.write(f"- Telegram summary image: `{image_path}`\n")
        f.write(f"- Excel report: `{report_path}`\n\n")
        dash = dashboard_url()
        if dash:
            f.write(f"- Dashboard: {dash}\n\n")
        f.write(message.replace("\n", "\n\n"))


def send_telegram(message: str, image_path: Optional[Path] = None) -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.getenv("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        print("[Telegram] TELEGRAM_BOT_TOKEN 또는 TELEGRAM_CHAT_ID가 없어 발송을 건너뜁니다.")
        return

    dash = dashboard_url()

    # 1) 요약 이미지 먼저 발송
    if image_path and image_path.exists():
        caption = "📊 이번 주 소싱 TOP 요약"
        if dash:
            caption += f"\n\n🔗 대시보드: {dash}"
        photo_url = f"https://api.telegram.org/bot{token}/sendPhoto"
        with image_path.open("rb") as f:
            r = requests.post(
                photo_url,
                data={"chat_id": chat_id, "caption": caption[:1000]},
                files={"photo": f},
                timeout=30,
            )
        r.raise_for_status()

    # 2) 텍스트 브리핑 발송
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    chunks = [message[i : i + 3500] for i in range(0, len(message), 3500)]
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


def send_email(message: str, report_path: Path, image_path: Optional[Path] = None) -> None:
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

    data = report_path.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=report_path.name,
    )

    if image_path and image_path.exists():
        img_data = image_path.read_bytes()
        msg.add_attachment(
            img_data,
            maintype="image",
            subtype="png",
            filename=image_path.name,
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
    parser.add_argument("--image-top", type=int, default=10)
    parser.add_argument("--no-image", action="store_true", help="텔레그램 요약 이미지 생성을 끕니다.")
    args = parser.parse_args()

    report_path = find_latest_report(args.reports_dir)
    message = build_message(report_path, top_n=args.top)

    image_path: Optional[Path] = None
    send_image_env = os.getenv("SEND_TELEGRAM_IMAGE", "true").strip().lower()
    if not args.no_image and send_image_env not in {"0", "false", "no", "off"}:
        image_path = create_summary_image(report_path, top_n=args.image_top, output_dir=args.reports_dir)

    write_github_summary(message, report_path, image_path=image_path)

    errors = []
    try:
        send_telegram(message, image_path=image_path)
    except Exception as e:
        errors.append(f"send_telegram: {e}")

    try:
        send_webhook(message)
    except Exception as e:
        errors.append(f"send_webhook: {e}")

    try:
        send_email(message, report_path, image_path=image_path)
    except Exception as e:
        errors.append(f"send_email: {e}")

    print(message)
    if image_path:
        print(f"\nTelegram summary image: {image_path}")

    if errors:
        print("\n알림 발송 중 일부 실패:")
        for e in errors:
            print("-", e)


if __name__ == "__main__":
    main()
