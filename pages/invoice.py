import io
import re
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

# 다중 키 매핑 사전 (수량 필드 추가)
_COLUMN_MAP = {
    "주문번호": ["Order ID", "주문 ID", "주문번호", "No", "Order Number"],
    "받는사람": ["Receiver name", "Receiver Name", "수취인 이름", "받는사람", "수령인", "이름","수령인 이름"],
    "전화번호1": ["Mobile", "모바일", "전화번호", "연락처", "휴대폰","수령인 전화번호"],
    "우편번호": ["Zip Code", "우편번호", "Zip", "POST","배송 우편번호(다음 우편번호로 발송해야 합니다.)"],
    "상세주소": ["Detailed address", "상세 주소", "상세주소", "주소2", "배송지","배송 주소 1"],
    "상품명": ["Product Information", "제품 정보", "상품명", "품명", "Item","선택 사항","구매 수량"],
    "수량1": ["발송할 수량", "수량", "주문수량", "Quantity", "수량(개)"] # 💡 수량 매핑 키워드 추가
}

_CITY_CANDIDATES = ["City", "city", "도시", "시", "시/군/구", "주소1"]

def _convert_auto(src: pd.DataFrame) -> pd.DataFrame:
    """다중 키를 순차적으로 탐색하여 데이터를 추출하고 포맷을 보정합니다."""
    
    out = pd.DataFrame(columns=_TARGET_COLUMNS)

    # 유연한 추출 도우미 함수
    def get_val(key_list):
        for k in key_list:
            if k in src.columns:
                return src[k].fillna("").astype(str).str.strip()
        return pd.Series([""] * len(src))

    # 1. 기본 데이터 매핑
    out["주문번호"] = get_val(_COLUMN_MAP["주문번호"])
    out["받는사람"] = get_val(_COLUMN_MAP["받는사람"])
    
    # 2. 전화번호 처리 (0010 방어 로직 적용)
    phone_series = get_val(_COLUMN_MAP["전화번호1"])
    phone_series = phone_series.str.replace(r'^\+82[-\s]*0*', '0', regex=True)
    out["전화번호1"] = phone_series
    out["전화번호2"] = phone_series  # 2번은 1번을 복사
    
    out["우편번호"] = get_val(_COLUMN_MAP["우편번호"])
    
    # 3. 주소 합성 로직 (테무/글로벌 포맷 우선 적용)
    target_addr_cols = ["배송 주", "배송 도시", "배송 주소 1", "배송 주소 2"]
    
    # 해당 4개의 컬럼 중 하나라도 파일에 존재하면 4단 병합 실행
    if any(c in src.columns for c in target_addr_cols):
        addr_parts = []
        for col in target_addr_cols:
            if col in src.columns:
                addr_parts.append(src[col].fillna("").astype(str).str.strip())
            else:
                addr_parts.append(pd.Series([""] * len(src)))
        
        # 파트들을 공백으로 연결 후, 연속된 다중 공백은 하나로 축소
        combined_addr = addr_parts[0] + " " + addr_parts[1] + " " + addr_parts[2] + " " + addr_parts[3]
        out["주소"] = combined_addr.str.replace(r'\s+', ' ', regex=True).str.strip()
    
    # 존재하지 않으면 기존 (도시 + 상세주소) 병합 실행
    else:
        city_col = next((c for c in _CITY_CANDIDATES if c in src.columns), None)
        city_part = src[city_col].fillna("").astype(str).str.strip() if city_col else pd.Series([""] * len(src))
        addr_part = get_val(_COLUMN_MAP["상세주소"]).str.strip()
        out["주소"] = (city_part + " " + addr_part).str.strip()
    
    # 4. 기타 필드 고정 및 수량 처리
    out["상품명1"] = get_val(_COLUMN_MAP["상품명"])
    
    # 💡 수량 데이터 가져오기 (비어있거나 매핑 컬럼이 없으면 '1'로 대체)
    qty_series = get_val(_COLUMN_MAP["수량1"])
    out["수량1"] = qty_series.replace("", "1") 

    return out[_TARGET_COLUMNS]

# ── 메인 UI ────────────────
st.title("📦 택배 송장 지능형 변환")
st.info("💡 테무(CSV) 및 스마트스토어/쿠팡(Excel) 발주서를 자동으로 인식하여 표준 양식으로 변환합니다.")

input_file = st.file_uploader("원본 주문 엑셀/CSV 선택", type=["xlsx", "xls", "csv"])

if input_file:
    if st.button("🚀 변환 실행"):
        try:
            # 파일 확장자에 따른 분기 처리 (CSV vs Excel)
            file_name = input_file.name.lower()

            if file_name.endswith(".csv"):
                # CSV 파일 처리 (UTF-8, CP949 인코딩 모두 방어)
                try:
                    src_df = pd.read_csv(input_file, dtype=str, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    input_file.seek(0)
                    src_df = pd.read_csv(input_file, dtype=str, encoding="cp949")
            else:
                # 엑셀 파일 처리
                src_df = pd.read_excel(input_file, dtype=str, engine="openpyxl" if file_name.endswith(".xlsx") else None)

            # 컬럼명 공백 제거로 매칭률 극대화
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
            
            st.subheader("👀 변환 결과 미리보기 (앞 5건)")
            st.dataframe(final_df.head(), use_container_width=True)

        except Exception as e:
            st.error(f"❌ 변환 오류: {e}")
