"""
pages/invoice.py — 📦 택배 송장 변환
원본 주문 엑셀을 요정비닐 A-type 템플릿 양식으로 변환합니다.
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st


# ── 컬럼 매핑 설정 ────────────────────────────────────────
_COLUMN_MAP = {
    "주문번호"  : "Order ID",
    "받는사람"  : "Receiver Name",
    "전화번호1" : "Mobile",
    "전화번호2" : "Mobile",       # 전화번호2 = 전화번호1과 동일
    "우편번호"  : "Zip Code",
    "주소"      : "Detailed address",
    "상품명1"   : "Product Information",
}

# 도시 정보 컬럼 후보 (순서대로 탐색)
_CITY_CANDIDATES = ["City", "city", "도시", "시", "시/군/구"]


def _convert(src: pd.DataFrame) -> pd.DataFrame:
    """원본 DataFrame을 요정비닐 양식으로 변환합니다."""
    city_col = next((c for c in _CITY_CANDIDATES if c in src.columns), None)
    if city_col is None:
        raise ValueError("원본 파일에서 'City' 또는 '도시' 컬럼을 찾을 수 없습니다.")

    out = pd.DataFrame()
    for out_col, src_col in _COLUMN_MAP.items():
        if src_col in src.columns:
            out[out_col] = src[src_col].fillna("").astype(str)

    # 주소 = 도시 + 상세주소
    out["주소"] = (
        src[city_col].fillna("").astype(str).str.strip()
        + " "
        + src["Detailed address"].fillna("").astype(str).str.strip()
    ).str.strip()

    # 전화번호2 = 전화번호1 복사
    out["전화번호2"] = out.get("전화번호1", "")

    return out


def render() -> None:
    st.title("📦 택배 송장 자동 변환기 (A-type)")
    st.write("원본 주문 엑셀을 요정비닐 템플릿 양식에 맞춰 변환합니다.")

    col1, col2 = st.columns(2)
    with col1:
        input_file = st.file_uploader("1. 원본 주문 엑셀 선택", type=["xlsx", "xls"])
    with col2:
        st.file_uploader("2. 템플릿 엑셀(A-type) 선택 (참고용)", type=["xlsx", "xls"])

    if not input_file:
        return

    if st.button("🚀 변환 실행"):
        try:
            src = pd.read_excel(input_file, dtype=str)
            out = _convert(src)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                out.to_excel(writer, index=False)

            st.success(f"✅ 변환 완료! (총 {len(out)}행)")
            st.download_button(
                label="📥 변환된 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"요정비닐_송장변환_{datetime.now().strftime('%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ValueError as ve:
            st.error(f"⚠️ {ve}")
        except Exception as e:
            st.error(f"❌ 변환 실패: {e}")
