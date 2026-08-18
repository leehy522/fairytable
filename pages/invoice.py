import io
from datetime import datetime
import pandas as pd
import streamlit as st
from auth import check_password

# 1. 페이지 설정 및 보안
st.set_page_config(page_title="택배 송장 변환", page_icon="📦", layout="wide")

if not check_password():
    st.stop()

# ── A-type 고정 출력 양식 ────────────────
_TARGET_COLUMNS = [
    "주문번호", "받는사람", "전화번호1", "전화번호2", "우편번호", 
    "주소", "상품명1", "수량1"
]

# 💡 귀하의 요청을 반영한 다중 키 매핑 사전 (리스트 형태)
_COLUMN_MAP = {
    "주문번호": ["Order ID", "주문 ID", "주문번호", "No", "Order Number"],
    "받는사람": ["Receiver name", "Receiver Name", "수취인 이름", "받는사람", "수령인", "이름","수령인 이름"],
    "전화번호1": ["Mobile", "모바일", "전화번호", "연락처", "휴대폰","수령인 전화번호"],
    "우편번호": ["Zip Code", "우편번호", "Zip", "POST","배송 우편번호(다음 우편번호로 발송해야 합니다.)"],
    "상세주소": ["Detailed address", "상세 주소", "상세주소", "주소2", "배송지","배송 주소 1"],
    "상품명": ["Product Information", "제품 정보", "상품명", "품명", "Item","선택 사항","구매 수량"]
}

_CITY_CANDIDATES = ["City", "city", "도시", "시", "시/군/구", "주소1"]

def _convert_auto(src: pd.DataFrame) -> pd.DataFrame:
    """다중 키를 순차적으로 탐색하여 데이터를 추출합니다."""
    
    city_col = next((c for c in _CITY_CANDIDATES if c in src.columns), None)
    out = pd.DataFrame(columns=_TARGET_COLUMNS)

    # 💡 유연한 추출 도우미 함수
    def get_val(key_list):
        for k in key_list:
            if k in src.columns:
                return src[k].fillna("").astype(str)
        return pd.Series([""] * len(src))

    # 데이터 매핑 실행
    out["주문번호"] = get_val(_COLUMN_MAP["주문번호"])
    out["받는사람"] = get_val(_COLUMN_MAP["받는사람"])
    out["전화번호1"] = get_val(_COLUMN_MAP["전화번호1"])
    out["전화번호2"] = out["전화번호1"] # 전화번호2는 1번을 복사
    out["우편번호"] = get_val(_COLUMN_MAP["우편번호"])
    
    # 주소 합성 (도시 + 상세주소)
    city_part = src[city_col].fillna("").astype(str).str.strip() if city_col else pd.Series([""] * len(src))
    addr_part = get_val(_COLUMN_MAP["상세주소"]).str.strip()
    out["주소"] = (city_part + " " + addr_part).str.strip()
    
    out["상품명1"] = get_val(_COLUMN_MAP["상품명"])
    out["수량1"] = "1" # 기본 수량 1 고정
    out["배송메시지"] = ""
    out["운송장번호"] = ""

    return out[_TARGET_COLUMNS]

# ── 메인 UI ────────────────
st.title("📦 택배 송장 지능형 변환")
st.info("💡 설정하신 한글/영문 컬럼명을 모두 검색하여 자동으로 변환합니다.")

input_file = st.file_uploader("원본 주문 엑셀 선택", type=["xlsx", "xls","csv"])

if input_file:
    if st.button("🚀 변환 실행"):
        try:
            # 엑셀 로드 시 컬럼명 공백 제거로 매칭률 극대화
            src_df = pd.read_excel(input_file, dtype=str, engine='openpyxl')
            src_df.columns = [str(c).strip() for c in src_df.columns]
            
            final_df = _convert_auto(src_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False)

            st.success(f"✅ 변환 완료! 총 {len(final_df)}건의 데이터를 처리했습니다.")
            st.download_button(
                label="📥 변환된 엑셀 다운로드",
                data=output.getvalue(),
                file_name=f"요정비닐_송장_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
            st.subheader("👀 변환 결과 미리보기")
            st.dataframe(final_df.head(), use_container_width=True)

        except Exception as e:
            st.error(f"❌ 변환 오류: {e}")
