import io
from datetime import datetime
import pandas as pd
import streamlit as st
from auth import check_password

# 1. 페이지 설정 및 보안
st.set_page_config(page_title="택배 송장 변환", page_icon="📦", layout="wide")

if not check_password():
    st.stop()

# ── A-type 고정 양식 설정 (업로드해주신 파일 기준) ────────────────
_TARGET_COLUMNS = [
    "주문번호", "받는사람", "전화번호1", "전화번호2", "우편번호", 
    "주소", "상품명1", "수량1", "배송메시지", "운송장번호"
]

# 원본에서 찾아야 할 매핑 키
_COLUMN_MAP = {
    "주문번호": "Order ID",
    "받는사람": "Receiver Name",
    "전화번호1": "Mobile",
    "전화번호2": "Mobile",
    "우편번호": "Zip Code",
    "상세주소": "Detailed address",
    "상품명": "Product Information"
}

_CITY_CANDIDATES = ["City", "city", "도시", "시", "시/군/구"]

def _convert_auto(src: pd.DataFrame) -> pd.DataFrame:
    """양식 파일 없이 원본을 A-type 고정 양식으로 즉시 변환합니다."""
    
    # 1. 도시 컬럼 탐색
    city_col = next((c for c in _CITY_CANDIDATES if c in src.columns), None)
    if city_col is None:
        raise ValueError("원본 파일에서 'City' 또는 '도시' 컬럼을 찾을 수 없습니다.")

    # 2. 고정 양식의 빈 틀 생성
    out = pd.DataFrame(columns=_TARGET_COLUMNS)

    # 3. 데이터 매핑 및 채우기
    out["주문번호"] = src.get(_COLUMN_MAP["주문번호"], "").fillna("").astype(str)
    out["받는사람"] = src.get(_COLUMN_MAP["받는사람"], "").fillna("").astype(str)
    out["전화번호1"] = src.get(_COLUMN_MAP["전화번호1"], "").fillna("").astype(str)
    out["전화번호2"] = out["전화번호1"]  # 전화번호2는 1번 복사
    out["우편번호"] = src.get(_COLUMN_MAP["우편번호"], "").fillna("").astype(str)
    
    # 주소 합성 (도시 + 상세주소)
    addr_detail = src.get(_COLUMN_MAP["상세주소"], "").fillna("").astype(str)
    out["주소"] = (src[city_col].fillna("").astype(str).str.strip() + " " + addr_detail.str.strip()).str.strip()
    
    out["상품명1"] = src.get(_COLUMN_MAP["상품명"], "").fillna("").astype(str)
    out["수량1"] = "1"  # 기본값 1 세팅
    out["배송메시지"] = "" # 필요 시 원본 매핑 가능
    out["운송장번호"] = "" # 비워둠

    return out[_TARGET_COLUMNS] # 정해진 순서대로 반환

# ── 메인 UI ──────────────────────────────────────────────
st.title("📦 택배 송장 자동 변환 (A-type 고정)")
st.info("💡 양식 파일을 올릴 필요가 없습니다. 원본 주문 엑셀만 업로드하세요.")

input_file = st.file_uploader("원본 주문 엑셀 선택", type=["xlsx", "xls"])

if input_file:
    if st.button("🚀 즉시 변환 및 다운로드"):
        try:
            # 원본 데이터 읽기
            src_df = pd.read_excel(input_file, dtype=str)
            
            # 자동 변환 실행
            final_df = _convert_auto(src_df)

            # 엑셀 파일 생성
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False)

            st.success(f"✅ 변환 완료! 요정비닐 A-type 양식으로 재구성되었습니다.")
            st.download_button(
                label="📥 변환된 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"요정비닐_송장_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
            # 미리보기 (상위 5개)
            st.subheader("👀 변환 결과 미리보기")
            st.table(final_df.head())

        except Exception as e:
            st.error(f"❌ 변환 중 오류 발생: {e}")
