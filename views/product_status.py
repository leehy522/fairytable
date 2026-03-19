"""
pages/product_status.py — 🏷️ 요정비닐 상품 현황
구글 스프레드시트를 CSV로 받아 실시간 표시합니다.
"""

import streamlit as st
import pandas as pd

# ── 구글 시트 CSV URL ─────────────────────────────────────
_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vTVvCbm9KEoUrqvlXSyIyLHmstIGZuiuTMLYDBnmgnxInrfoMelDXFSWogUdHUfNALb7uC_nBAIyzif"
    "/pub?output=csv"
)


@st.cache_data(ttl=60)
def _load_data() -> pd.DataFrame:
    """구글 시트를 1분 캐시로 읽어옵니다."""
    try:
        df = pd.read_csv(_CSV_URL)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(subset=["상품명"])
    except Exception as e:
        st.error(f"구글 시트 연결 오류: {e}")
        return pd.DataFrame()


def render() -> None:
    st.title("🏷️ 요정비닐 상품 실시간 현황")
    st.caption("구글 스프레드시트와 동기화 중 (자동 갱신 주기: 1분)")
    st.divider()

    df = _load_data()

    if df.empty:
        st.error(
            "데이터를 불러올 수 없습니다. 구글 시트의 '웹에 게시' 설정을 확인해주세요."
        )
        return

    col1, _ = st.columns([1, 1])
    with col1:
        st.metric("총 등록 상품", f"{len(df)} 종")

    with st.expander("🔍 상품 검색 및 필터", expanded=True):
        search_query = st.text_input("", placeholder="검색할 상품명을 입력하세요")

    display_df = (
        df[df["상품명"].str.contains(search_query, na=False)] if search_query else df
    )

    st.subheader(f"📦 상품 목록 ({len(display_df)}건)")
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "상품명": st.column_config.TextColumn("📋 상품명"),
            "단가": st.column_config.NumberColumn("💰 단가", format="₩%d"),
            "재고": st.column_config.NumberColumn("🔢 현재고", format="%d"),
        },
    )
    st.success("✅ 최신 데이터가 반영되었습니다.")