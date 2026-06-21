import io
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="네이버 해외직구 소싱 대시보드",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


DATA_FILE = Path("dashboard_data/latest_sourcing_rank.xlsx")
REPORTS_DIR = Path("reports")

PREFERRED_SHEETS = [
    "상품키워드_TOP",
    "신규발굴_TOP",
    "TOP_100",
    "TOP_80",
    "전체랭킹",
]

COLUMN_MAP = {
    "recommendation": "추천구분",
    "final_score": "소싱점수",
    "score": "소싱점수",
    "brand": "브랜드",
    "brand_name": "브랜드",
    "브랜드명": "브랜드",
    "category": "카테고리",
    "대표카테고리": "카테고리",
    "keyword": "키워드",
    "product_keyword": "키워드",
    "search_volume": "검색량",
    "monthly_search_volume": "검색량",
    "search_volume_growth_pct": "검색량증감률",
    "total_products": "전체상품수",
    "네이버 상품수": "전체상품수",
    "네이버\n상품수": "전체상품수",
    "overseas_products": "해외직구상품수",
    "네이버 해외 상품수": "해외직구상품수",
    "네이버 해외\n상품수": "해외직구상품수",
    "overseas_ratio": "해외직구비중",
    "avg_top10_price": "평균가",
    "국내 평균가": "평균가",
    "min_top10_price": "최저가",
    "국내 최저가": "최저가",
    "risk_flags": "리스크",
    "risk_penalty": "리스크감점",
    "shopping_status": "쇼핑API상태",
    "sample_titles": "샘플상품명",
    "source_type": "후보출처",
}

NUMERIC_COLS = [
    "소싱점수",
    "검색량",
    "검색량증감률",
    "전체상품수",
    "해외직구상품수",
    "해외직구비중",
    "평균가",
    "최저가",
    "리스크감점",
]

TEXT_COLS = [
    "추천구분",
    "브랜드",
    "카테고리",
    "키워드",
    "리스크",
    "쇼핑API상태",
    "샘플상품명",
    "후보출처",
]


def latest_report_path() -> Path | None:
    if DATA_FILE.exists():
        return DATA_FILE
    if REPORTS_DIR.exists():
        files = sorted(REPORTS_DIR.glob("sourcing_rank_*.xlsx"), reverse=True)
        if files:
            return files[0]
    return None


@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    buffer = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(buffer, engine="openpyxl")
    return {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}


@st.cache_data(show_spinner=False)
def read_excel_path(path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path, engine="openpyxl")
    return {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={c: COLUMN_MAP.get(c, c) for c in df.columns})

    for c in TEXT_COLS:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).replace({"nan": "", "None": ""})

    for c in NUMERIC_COLS:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    if "키워드" not in df.columns or df["키워드"].eq("").all():
        df["키워드"] = df["브랜드"]

    df["상품키워드여부"] = df.apply(
        lambda r: is_product_keyword(r.get("브랜드", ""), r.get("키워드", "")),
        axis=1,
    )

    df["API오류"] = (
        df["쇼핑API상태"].str.contains("error|fail|timeout|414|too large", case=False, na=False)
        | ((df["전체상품수"] <= 0) & (df["평균가"] <= 0))
    )

    if "해외직구비중" in df.columns:
        ratio = df["해외직구상품수"] / df["전체상품수"].replace(0, pd.NA)
        df["해외직구비중"] = pd.to_numeric(df["해외직구비중"], errors="coerce").fillna(ratio.fillna(0)).fillna(0)

    return df


def is_product_keyword(brand: str, keyword: str) -> bool:
    brand = str(brand).strip().lower()
    keyword = str(keyword).strip().lower()
    if not keyword:
        return False
    if brand and keyword == brand:
        return False
    if " " in keyword:
        return True
    if brand and brand in keyword and len(keyword) >= len(brand) + 3:
        return True
    return False


def pick_sheet(sheet_names: list[str]) -> str:
    for s in PREFERRED_SHEETS:
        if s in sheet_names:
            return s
    return sheet_names[0]


def comma_int(v) -> str:
    try:
        return f"{int(float(v)):,}"
    except Exception:
        return "0"


def price(v) -> str:
    try:
        n = int(float(v))
        return "-" if n <= 0 else f"{n:,}원"
    except Exception:
        return "-"


def make_title(row: pd.Series) -> str:
    brand = str(row.get("브랜드", "")).strip()
    keyword = str(row.get("키워드", "")).strip()
    title = keyword if brand.lower() in keyword.lower() else f"{brand} {keyword}".strip()
    title = re.sub(r"(무료배송|당일발송|특가|최저가|인기추천|정품보장)", "", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title[:80]


def make_tags(row: pd.Series) -> str:
    raw = " ".join([str(row.get("브랜드", "")), str(row.get("카테고리", "")), str(row.get("키워드", ""))])
    tags, seen = [], set()
    for x in re.split(r"[\s,/|]+", raw):
        x = x.strip("#[]()·")
        if 1 < len(x) <= 20 and x not in seen:
            tags.append(x)
            seen.add(x)
        if len(tags) >= 10:
            break
    return ", ".join(tags)


def load_data_from_sidebar() -> tuple[dict[str, pd.DataFrame], str]:
    with st.sidebar:
        st.header("데이터")
        uploaded = st.file_uploader("엑셀 리포트 업로드", type=["xlsx"])

    if uploaded is not None:
        return read_excel_bytes(uploaded.getvalue()), uploaded.name

    path = latest_report_path()
    if path:
        return read_excel_path(str(path)), str(path)

    return {}, ""


def render_dashboard():
    st.title("📦 네이버 해외직구 상품키워드 소싱 대시보드")
    st.caption("브랜드가 아니라 실제 등록 가능한 상품/모델 키워드를 빠르게 선별하는 화면입니다.")

    sheets, source_label = load_data_from_sidebar()

    if not sheets:
        st.info("`dashboard_data/latest_sourcing_rank.xlsx`가 아직 없거나, 읽을 엑셀 파일이 없습니다. 왼쪽에서 리포트 엑셀을 업로드하세요.")
        return

    sheet_names = list(sheets.keys())
    initial = pick_sheet(sheet_names)

    with st.sidebar:
        sheet = st.selectbox("시트 선택", sheet_names, index=sheet_names.index(initial))
        st.divider()
        st.header("필터")

    df = normalize_df(sheets[sheet])

    with st.sidebar:
        product_only = st.checkbox("상품/모델 키워드만 보기", value=("상품" in sheet))
        hide_api_errors = st.checkbox("API 오류/미조회 제외", value=False)
        exclude_risk = st.checkbox("리스크 키워드 제외", value=False)

        min_score = st.slider("최소 소싱점수", 0.0, 100.0, 0.0, step=1.0)

        max_volume = int(max(1000, df["검색량"].max() if len(df) else 1000))
        min_volume = st.slider("최소 검색량", 0, max_volume, 0, step=max(1, max_volume // 100))

        max_price = int(max(100000, df["평균가"].max() if len(df) else 100000))
        min_price = st.slider("최소 평균가", 0, max_price, 0, step=max(1000, max_price // 100))

        q = st.text_input("브랜드/키워드 검색", "")

    filtered = df.copy()
    filtered = filtered[(filtered["소싱점수"] >= min_score) & (filtered["검색량"] >= min_volume) & (filtered["평균가"] >= min_price)]

    if product_only:
        filtered = filtered[filtered["상품키워드여부"]]
    if hide_api_errors:
        filtered = filtered[~filtered["API오류"]]
    if exclude_risk:
        risk_pattern = "식품|화장품|의약|건강|전기|전자|KC|파손|부피|상표|짝퉁"
        filtered = filtered[~filtered["리스크"].str.contains(risk_pattern, case=False, na=False)]
    if q.strip():
        pattern = re.escape(q.strip())
        filtered = filtered[
            filtered["브랜드"].str.contains(pattern, case=False, na=False)
            | filtered["키워드"].str.contains(pattern, case=False, na=False)
            | filtered["카테고리"].str.contains(pattern, case=False, na=False)
        ]

    st.write(f"**데이터:** `{source_label}` · **시트:** `{sheet}`")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("후보 수", comma_int(len(filtered)))
    c2.metric("평균 점수", f"{filtered['소싱점수'].mean():.1f}" if len(filtered) else "0")
    c3.metric("상품키워드", comma_int(filtered["상품키워드여부"].sum()))
    c4.metric("API 오류", comma_int(filtered["API오류"].sum()))
    c5.metric("검색량 합계", comma_int(filtered["검색량"].sum()))

    tab1, tab2, tab3 = st.tabs(["TOP 리스트", "상세 검토", "상품명/태그 초안"])

    display_cols = [
        "추천구분", "소싱점수", "브랜드", "카테고리", "키워드", "검색량",
        "전체상품수", "해외직구상품수", "해외직구비중", "평균가", "최저가",
        "리스크", "후보출처", "API오류",
    ]
    display_cols = [c for c in display_cols if c in filtered.columns]
    sorted_df = filtered.sort_values("소싱점수", ascending=False)

    with tab1:
        st.dataframe(sorted_df[display_cols], use_container_width=True, hide_index=True)
        st.download_button(
            "필터 결과 CSV 다운로드",
            sorted_df[display_cols].to_csv(index=False).encode("utf-8-sig"),
            file_name=f"sourcing_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )

    with tab2:
        top = sorted_df.head(20)
        if len(top):
            for _, row in top.iterrows():
                with st.expander(f"{row.get('키워드', '')} · {float(row.get('소싱점수', 0)):.1f}점"):
                    a, b, c, d = st.columns(4)
                    a.metric("브랜드", str(row.get("브랜드", "")))
                    b.metric("검색량", comma_int(row.get("검색량", 0)))
                    c.metric("평균가", price(row.get("평균가", 0)))
                    d.metric("해외직구 상품수", comma_int(row.get("해외직구상품수", 0)))
                    st.write("**카테고리:**", row.get("카테고리", ""))
                    st.write("**리스크:**", row.get("리스크", "") or "-")
                    if str(row.get("샘플상품명", "")).strip():
                        st.write("**샘플 상품명:**", row.get("샘플상품명", ""))
        else:
            st.info("표시할 데이터가 없습니다.")

    with tab3:
        top = sorted_df.head(30)
        if len(top):
            title_df = pd.DataFrame(
                {
                    "브랜드": top["브랜드"],
                    "키워드": top["키워드"],
                    "상품명 초안": [make_title(r) for _, r in top.iterrows()],
                    "태그 후보": [make_tags(r) for _, r in top.iterrows()],
                    "소싱점수": top["소싱점수"],
                    "평균가": top["평균가"],
                }
            )
            st.dataframe(title_df, use_container_width=True, hide_index=True)
            st.download_button(
                "상품명/태그 초안 CSV 다운로드",
                title_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"smartstore_titles_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
            )
        else:
            st.info("표시할 데이터가 없습니다.")


def main():
    try:
        render_dashboard()
    except Exception as e:
        st.error("대시보드 실행 중 오류가 발생했습니다.")
        st.exception(e)
        st.info("오류 로그를 복사해서 보내주시면 바로 수정할 수 있습니다. 임시로는 왼쪽에서 최신 리포트 엑셀을 다시 업로드해보세요.")


if __name__ == "__main__":
    main()
