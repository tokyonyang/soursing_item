import json
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


DATA_CANDIDATES = [
    Path("dashboard_data/latest_sourcing_rank.xlsx"),
    Path("reports/latest_sourcing_rank.xlsx"),
]

PREFERRED_SHEETS = [
    "상품키워드_TOP",
    "신규발굴_TOP",
    "TOP_100",
    "TOP_80",
    "전체랭킹",
]


COLUMN_MAP = {
    # English/internal columns
    "recommendation": "추천구분",
    "final_score": "소싱점수",
    "brand": "브랜드",
    "category": "카테고리",
    "keyword": "키워드",
    "search_volume": "검색량",
    "prev_search_volume": "이전검색량",
    "search_volume_growth_pct": "검색량증감률",
    "total_products": "전체상품수",
    "overseas_products": "해외직구상품수",
    "overseas_ratio": "해외직구비중",
    "opportunity_raw": "기회지수",
    "avg_top10_price": "평균가",
    "min_top10_price": "최저가",
    "risk_flags": "리스크",
    "risk_penalty": "리스크감점",
    "volume_source": "검색량출처",
    "shopping_status": "쇼핑API상태",
    "sample_titles": "샘플상품명",
    "sample_malls": "샘플몰",
    "source_type": "후보출처",

    # Korean-ish variants
    "브랜드명": "브랜드",
    "대표카테고리": "카테고리",
    "네이버 상품수": "전체상품수",
    "네이버\n상품수": "전체상품수",
    "네이버 해외 상품수": "해외직구상품수",
    "네이버 해외\n상품수": "해외직구상품수",
    "국내 평균가": "평균가",
    "국내 최저가": "최저가",
    "검색량 증가율": "검색량증감률",
    "추천": "추천구분",
    "점수": "소싱점수",
}


NUMERIC_COLUMNS = [
    "소싱점수",
    "검색량",
    "이전검색량",
    "검색량증감률",
    "전체상품수",
    "해외직구상품수",
    "해외직구비중",
    "기회지수",
    "평균가",
    "최저가",
    "리스크감점",
]


RISK_KEYWORDS = [
    "식품",
    "화장품",
    "의약",
    "건강",
    "영양",
    "전기",
    "전자",
    "KC",
    "사이즈",
    "파손",
    "부피",
    "정품증빙",
    "짝퉁",
    "상표",
]


def find_default_data_file() -> Path | None:
    for p in DATA_CANDIDATES:
        if p.exists():
            return p

    reports = sorted(Path("reports").glob("sourcing_rank_*.xlsx"), reverse=True)
    if reports:
        return reports[0]
    return None


@st.cache_data(show_spinner=False)
def load_workbook_from_bytes(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(file_bytes)
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}


@st.cache_data(show_spinner=False)
def load_workbook_from_path(path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    renamed = {}
    for col in df.columns:
        key = col.strip()
        renamed[col] = COLUMN_MAP.get(key, key)
    df = df.rename(columns=renamed)

    for col in [
        "추천구분",
        "브랜드",
        "카테고리",
        "키워드",
        "리스크",
        "검색량출처",
        "쇼핑API상태",
        "샘플상품명",
        "샘플몰",
        "후보출처",
    ]:
        if col not in df.columns:
            df[col] = ""

    for col in NUMERIC_COLUMNS:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["브랜드"] = df["브랜드"].astype(str).replace("nan", "")
    df["키워드"] = df["키워드"].astype(str).replace("nan", "")
    df["카테고리"] = df["카테고리"].astype(str).replace("nan", "")
    df["리스크"] = df["리스크"].astype(str).replace("nan", "")
    df["쇼핑API상태"] = df["쇼핑API상태"].astype(str).replace("nan", "")
    df["후보출처"] = df["후보출처"].astype(str).replace({"": "seed", "nan": "seed"})

    df["API오류"] = (
        df["쇼핑API상태"].str.contains("error|414|Too Large|timeout|fail", case=False, na=False)
        | ((df["전체상품수"] <= 0) & (df["평균가"] <= 0))
    )

    df["상품키워드여부"] = df.apply(
        lambda r: is_product_keyword(r.get("브랜드", ""), r.get("키워드", "")),
        axis=1,
    )

    df["가격대"] = pd.cut(
        df["평균가"],
        bins=[-1, 0, 30000, 70000, 150000, 300000, 999999999],
        labels=["가격없음", "3만원↓", "3~7만원", "7~15만원", "15~30만원", "30만원↑"],
    ).astype(str)

    df["해외직구비중"] = df["해외직구비중"].where(
        df["해외직구비중"] > 0,
        (df["해외직구상품수"] / df["전체상품수"].replace(0, pd.NA)).fillna(0),
    )

    df["상품수대비검색량"] = df["검색량"] / (df["전체상품수"].replace(0, pd.NA))
    df["상품수대비검색량"] = pd.to_numeric(df["상품수대비검색량"], errors="coerce").fillna(0)

    return df


def is_product_keyword(brand: str, keyword: str) -> bool:
    brand = str(brand).strip().lower()
    keyword = str(keyword).strip().lower()
    if not keyword:
        return False
    if brand and keyword == brand:
        return False
    # 브랜드명만 있는 경우보다, 품목/모델 단어가 붙은 경우를 상품키워드로 간주
    if brand and keyword.startswith(brand) and len(keyword.replace(brand, "").strip()) >= 2:
        return True
    if " " in keyword and len(keyword) > len(brand) + 2:
        return True
    # 한글 키워드는 공백 없이 붙는 경우가 많아 길이 조건을 추가
    if brand and brand in keyword and len(keyword) >= len(brand) + 3:
        return True
    return False


def pick_initial_sheet(sheets: list[str]) -> str:
    for s in PREFERRED_SHEETS:
        if s in sheets:
            return s
    return sheets[0]


def comma_int(value) -> str:
    try:
        return f"{int(float(value)):,}"
    except Exception:
        return "0"


def comma_price(value) -> str:
    try:
        v = int(float(value))
        if v <= 0:
            return "-"
        return f"{v:,}원"
    except Exception:
        return "-"


def make_smartstore_title(row: pd.Series) -> str:
    brand = str(row.get("브랜드", "")).strip()
    keyword = str(row.get("키워드", "")).strip()

    # 너무 긴 키워드는 앞 45자만 사용
    keyword = re.sub(r"\s+", " ", keyword)
    if len(keyword) > 55:
        keyword = keyword[:55].strip()

    if brand and keyword and brand.lower() not in keyword.lower():
        title = f"{brand} {keyword}"
    else:
        title = keyword or brand

    # 스마트스토어 상품명에 불리한 과한 수식어 제거
    title = re.sub(r"(무료배송|당일발송|특가|최저가|정품보장|인기추천)", "", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title[:80]


def make_tags(row: pd.Series) -> list[str]:
    brand = str(row.get("브랜드", "")).strip()
    category = str(row.get("카테고리", "")).strip()
    keyword = str(row.get("키워드", "")).strip()

    raw = []
    for token in [brand, category, keyword]:
        for piece in re.split(r"[\s,/|]+", token):
            piece = piece.strip("#[]()·")
            if 1 < len(piece) <= 20:
                raw.append(piece)

    # 중복 제거, 최대 10개
    seen = set()
    tags = []
    for t in raw:
        if t not in seen:
            tags.append(t)
            seen.add(t)
        if len(tags) >= 10:
            break
    return tags


def render_metric_cards(df: pd.DataFrame):
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("후보 수", comma_int(len(df)))
    c2.metric("평균 점수", f"{df['소싱점수'].mean():.1f}" if len(df) else "0")
    c3.metric("상품키워드", comma_int(df["상품키워드여부"].sum()))
    c4.metric("신규발굴", comma_int((df["후보출처"].str.lower() == "discovered").sum()))
    c5.metric("API 오류", comma_int(df["API오류"].sum()))
    c6.metric("검색량 합계", comma_int(df["검색량"].sum()))


def render_data_quality(df: pd.DataFrame):
    if not len(df):
        return

    error_rate = df["API오류"].mean()
    if error_rate >= 0.3:
        st.warning(
            f"API 오류/미조회 비율이 {error_rate:.0%}입니다. 상품수·가격 데이터가 비어 있으면 랭킹 신뢰도가 떨어질 수 있습니다. "
            "검색어가 너무 길거나 API 키/호출 제한을 확인하세요."
        )

    long_keywords = df["키워드"].astype(str).str.len().gt(80).sum()
    if long_keywords:
        st.info(
            f"80자 이상 긴 키워드가 {long_keywords:,}개 있습니다. 브랜드별 대표 상품키워드 단위로 줄여 조회하는 패치를 적용하는 것이 좋습니다."
        )


def main():
    st.title("📦 네이버 해외직구 상품키워드 소싱 대시보드")
    st.caption("브랜드가 아니라 실제 등록 가능한 상품/모델 키워드를 빠르게 선별하는 화면입니다.")

    default_file = find_default_data_file()

    with st.sidebar:
        st.header("데이터")
        uploaded = st.file_uploader("엑셀 리포트 업로드", type=["xlsx"])

        if uploaded is not None:
            sheets = load_workbook_from_bytes(uploaded.getvalue())
            data_source_label = uploaded.name
        elif default_file:
            sheets = load_workbook_from_path(str(default_file))
            data_source_label = str(default_file)
        else:
            sheets = {}

        if not sheets:
            st.info("`dashboard_data/latest_sourcing_rank.xlsx`를 추가하거나, 리포트 엑셀을 업로드하세요.")
            st.stop()

        sheet_names = list(sheets.keys())
        initial_sheet = pick_initial_sheet(sheet_names)
        sheet = st.selectbox("시트 선택", sheet_names, index=sheet_names.index(initial_sheet))

    raw_df = sheets[sheet]
    df = normalize_columns(raw_df)

    st.write(f"**데이터:** `{data_source_label}` · **시트:** `{sheet}`")
    render_data_quality(df)

    with st.sidebar:
        st.header("필터")

        product_only = st.checkbox("상품/모델 키워드만 보기", value=("상품" in sheet or "키워드" in sheet))
        hide_api_errors = st.checkbox("API 오류/미조회 제외", value=False)
        exclude_risk = st.checkbox("리스크 키워드 제외", value=False)

        min_score, max_score = st.slider(
            "소싱점수",
            0.0,
            100.0,
            (float(max(0, min(50, df["소싱점수"].min() if len(df) else 0))), 100.0),
            step=1.0,
        )

        max_volume_default = int(max(1000, df["검색량"].max() if len(df) else 1000))
        volume_range = st.slider(
            "검색량",
            0,
            max_volume_default,
            (0, max_volume_default),
            step=max(1, max_volume_default // 100),
        )

        price_max = int(max(100000, df["평균가"].max() if len(df) else 100000))
        price_range = st.slider(
            "평균가",
            0,
            price_max,
            (0, price_max),
            step=max(1000, price_max // 100),
        )

        categories = sorted([c for c in df["카테고리"].dropna().unique().tolist() if c and c != "nan"])
        selected_categories = st.multiselect("카테고리", categories, default=[])

        recs = sorted([c for c in df["추천구분"].dropna().unique().tolist() if c and c != "nan"])
        selected_recs = st.multiselect("추천구분", recs, default=[])

        keyword_search = st.text_input("브랜드/키워드 검색", "")

    filtered = df.copy()

    filtered = filtered[
        (filtered["소싱점수"] >= min_score)
        & (filtered["소싱점수"] <= max_score)
        & (filtered["검색량"] >= volume_range[0])
        & (filtered["검색량"] <= volume_range[1])
        & (filtered["평균가"] >= price_range[0])
        & (filtered["평균가"] <= price_range[1])
    ]

    if product_only:
        filtered = filtered[filtered["상품키워드여부"]]
    if hide_api_errors:
        filtered = filtered[~filtered["API오류"]]
    if selected_categories:
        filtered = filtered[filtered["카테고리"].isin(selected_categories)]
    if selected_recs:
        filtered = filtered[filtered["추천구분"].isin(selected_recs)]
    if exclude_risk:
        pattern = "|".join(map(re.escape, RISK_KEYWORDS))
        filtered = filtered[~filtered["리스크"].str.contains(pattern, case=False, na=False)]
    if keyword_search.strip():
        q = keyword_search.strip()
        filtered = filtered[
            filtered["브랜드"].str.contains(q, case=False, na=False)
            | filtered["키워드"].str.contains(q, case=False, na=False)
            | filtered["카테고리"].str.contains(q, case=False, na=False)
        ]

    st.subheader("요약")
    render_metric_cards(filtered)

    tab1, tab2, tab3, tab4 = st.tabs(["TOP 리스트", "차트", "상세 검토", "상품명/태그 초안"])

    display_cols = [
        "추천구분",
        "소싱점수",
        "브랜드",
        "카테고리",
        "키워드",
        "검색량",
        "검색량증감률",
        "전체상품수",
        "해외직구상품수",
        "해외직구비중",
        "평균가",
        "최저가",
        "리스크",
        "후보출처",
        "API오류",
    ]
    display_cols = [c for c in display_cols if c in filtered.columns]

    with tab1:
        st.markdown("#### 필터 적용 결과")
        st.dataframe(
            filtered[display_cols].sort_values("소싱점수", ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "소싱점수": st.column_config.NumberColumn(format="%.1f"),
                "검색량": st.column_config.NumberColumn(format="%d"),
                "전체상품수": st.column_config.NumberColumn(format="%d"),
                "해외직구상품수": st.column_config.NumberColumn(format="%d"),
                "해외직구비중": st.column_config.NumberColumn(format="%.1%"),
                "평균가": st.column_config.NumberColumn(format="%d원"),
                "최저가": st.column_config.NumberColumn(format="%d원"),
                "API오류": st.column_config.CheckboxColumn(),
            },
        )

        csv = filtered[display_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "필터 결과 CSV 다운로드",
            csv,
            file_name=f"sourcing_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )

    with tab2:
        st.markdown("#### 점수 TOP 30")
        top_chart = filtered.sort_values("소싱점수", ascending=False).head(30)
        if len(top_chart):
            st.bar_chart(top_chart.set_index("키워드")["소싱점수"])
        else:
            st.info("표시할 데이터가 없습니다.")

        st.markdown("#### 검색량 대비 경쟁 확인")
        scatter_cols = ["검색량", "전체상품수", "소싱점수", "키워드"]
        if all(c in filtered.columns for c in scatter_cols) and len(filtered):
            # Streamlit scatter_chart는 컬럼명 기반 간단 시각화에 적합
            scatter_df = filtered[["키워드", "검색량", "전체상품수", "소싱점수"]].copy()
            st.scatter_chart(scatter_df, x="전체상품수", y="검색량", size="소싱점수")
        else:
            st.info("산점도에 필요한 컬럼이 없습니다.")

    with tab3:
        st.markdown("#### 상위 후보 상세 검토")
        top_detail = filtered.sort_values("소싱점수", ascending=False).head(20)
        if not len(top_detail):
            st.info("표시할 데이터가 없습니다.")
        for idx, row in top_detail.iterrows():
            title = f"{row.get('키워드', '')} · {row.get('소싱점수', 0):.1f}점"
            with st.expander(title):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("브랜드", str(row.get("브랜드", "")))
                c2.metric("검색량", comma_int(row.get("검색량", 0)))
                c3.metric("평균가", comma_price(row.get("평균가", 0)))
                c4.metric("해외직구 상품수", comma_int(row.get("해외직구상품수", 0)))

                st.write("**카테고리:**", row.get("카테고리", ""))
                st.write("**리스크:**", row.get("리스크", "") or "-")
                st.write("**API 상태:**", row.get("쇼핑API상태", "") or "-")

                sample_titles = str(row.get("샘플상품명", "")).strip()
                if sample_titles and sample_titles != "nan":
                    st.write("**샘플 상품명:**")
                    st.write(sample_titles)

    with tab4:
        st.markdown("#### 스마트스토어 등록 초안")
        top_titles = filtered.sort_values("소싱점수", ascending=False).head(30).copy()
        if not len(top_titles):
            st.info("표시할 데이터가 없습니다.")
        else:
            title_rows = []
            for _, row in top_titles.iterrows():
                title_rows.append(
                    {
                        "브랜드": row.get("브랜드", ""),
                        "키워드": row.get("키워드", ""),
                        "상품명 초안": make_smartstore_title(row),
                        "태그 후보": ", ".join(make_tags(row)),
                        "소싱점수": row.get("소싱점수", 0),
                        "평균가": row.get("평균가", 0),
                    }
                )
            title_df = pd.DataFrame(title_rows)
            st.dataframe(
                title_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "소싱점수": st.column_config.NumberColumn(format="%.1f"),
                    "평균가": st.column_config.NumberColumn(format="%d원"),
                },
            )

            st.download_button(
                "상품명/태그 초안 CSV 다운로드",
                title_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"smartstore_titles_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
            )


if __name__ == "__main__":
    main()
