import io
from datetime import datetime
import pandas as pd
import streamlit as st
from auth import check_password

# 1. 페이지 설정 및 보안
st.set_page_config(page_title="택배 송장 변환", page_icon="📦", layout="wide")

if not check_password():
    st.stop()

# ── A-type 고정 양식 설정 ────────────────
_TARGET_COLUMNS = [
    "주문번호", "받는사람", "전화번호1", "전화번호2", "우편번호", 
    "주소", "상품명1", "수량1", "배송메시지", "운송장번호"
]

_COLUMN_MAP = {
    "주문번호": "Order ID",
    "받는사람": "Receiver Name",
    "전화번호1": "Mobile",
    "우편번호": "Zip Code",
    "상세주소": "Detailed address",
    "상품명": "Product Information"
}

_CITY_CANDIDATES = ["City", "city", "도시", "시", "시/군/구"]

def _convert_auto(src: pd.DataFrame) -> pd.DataFrame:
    """오류를 수정한 안전한 데이터 매핑 로직입니다."""
    
    city_col = next((c for c in _CITY_CANDIDATES if c in src.columns), None)
    if city_col is None:
        raise ValueError("원본 파일에서 'City' 또는 '도시' 컬럼을 찾을 수 없습니다.")

    out = pd.DataFrame(columns=_TARGET_COLUMNS)

    # 안전한 데이터 추출 함수 (Series 객체 보존)
    def get_val(col_name):
        if col_name in src.columns:
            return src[col_name].fillna("").astype(str)
        return pd.Series([""] * len(src))

    # 데이터 매핑 (오류 수정 지점)
    out["주문번호"] = get_val(_COLUMN_MAP["주문번호"])
    out["받는사람"] = get_val(_COLUMN_MAP["받는사람"])
    out["전화번호1"] = get_val(_COLUMN_MAP["전화번호1"])
    out["전화번호2"] = out["전화번호1"]
    out["우편번호"] = get_val(_COLUMN_MAP["우편번호"])
    
    # 주소 합성
    city_part = src[city_col].fillna("").astype(str).str.strip()
    addr_part = get_val(_COLUMN_MAP["상세주소"]).str.strip()
    out["주소"] = (city_part + " " + addr_part).str.strip()
    
    out["상품명1"] = get_val(_COLUMN_MAP["상품명"])
    out["수량1"] = "1"
    out["배송메시지"] = ""
    out["운송장번호"] = ""

    return out[_TARGET_COLUMNS]

# ── 메인 UI ────────────────
st.title("📦 택배 송장 자동 변환 (A-type 고정)")

input_file = st.file_uploader("원본 주문 엑셀 선택", type=["xlsx", "xls"])

if input_file:
    if st.button("🚀 즉시 변환 및 다운로드"):
        try:
            # 엑셀 엔진을 openpyxl로 명시하여 안정성 확보
            src_df = pd.read_excel(input_file, dtype=str, engine='openpyxl')
            
            final_df = _convert_auto(src_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False)

            st.success(f"✅ 변환 완료! 요정비닐 A-type 양식으로 구성되었습니다.")
            st.download_button(
                label="📥 변환된 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"요정비닐_송장_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
            st.subheader("👀 변환 결과 미리보기")
            st.dataframe(final_df.head(), use_container_width=True)

        except Exception as e:
            st.error(f"❌ 변환 중 오류 발생: {e}")
