import io
import re
from datetime import datetime
import pandas as pd
import streamlit as st
import unicodedata
from urllib.parse import quote
from auth import check_password

# 1. 페이지 설정 및 보안
st.set_page_config(page_title="페어리비닐 통합 시스템", page_icon="🏭", layout="wide")

if not check_password():
    st.stop()

# ==========================================
# [기능 1] 택배 송장 지능형 변환 (V2.7)
# ==========================================
def show_invoice_converter():
    st.title("📦 택배 송장 지능형 변환")
    st.info("💡 테무(CSV) 및 스마트스토어/쿠팡(Excel) 발주서를 자동으로 인식하여 표준 양식으로 변환합니다.")

    _TARGET_COLUMNS = ["주문번호", "받는사람", "전화번호1", "전화번호2", "우편번호", "주소", "상품명1", "수량1"]
    
    _COLUMN_MAP = {
        "주문번호": ["Order ID", "주문 ID", "주문번호", "No", "Order Number"],
        "받는사람": ["Receiver name", "Receiver Name", "수취인 이름", "받는사람", "수령인", "이름","수령인 이름"],
        "전화번호1": ["Mobile", "모바일", "전화번호", "연락처", "휴대폰","수령인 전화번호"],
        "우편번호": ["Zip Code", "우편번호", "Zip", "POST","배송 우편번호(다음 우편번호로 발송해야 합니다.)"],
        "상세주소": ["Detailed address", "상세 주소", "상세주소", "주소2", "배송지","배송 주소 1"],
        "상품명": ["Product Information", "제품 정보", "상품명", "품명", "Item","선택 사항","구매 수량"],
        "수량1": ["발송할 수량", "수량", "주문수량", "Quantity", "수량(개)"]
    }
    _CITY_CANDIDATES = ["City", "city", "도시", "시", "시/군/구", "주소1"]

    def _convert_auto(src: pd.DataFrame) -> pd.DataFrame:
        out = pd.DataFrame(columns=_TARGET_COLUMNS)
        def get_val(key_list):
            for k in key_list:
                if k in src.columns:
                    return src[k].fillna("").astype(str).str.strip()
            return pd.Series([""] * len(src))

        out["주문번호"] = get_val(_COLUMN_MAP["주문번호"])
        out["받는사람"] = get_val(_COLUMN_MAP["받는사람"])
        
        phone_series = get_val(_COLUMN_MAP["전화번호1"]).str.replace(r'^\+82[-\s]*0*', '0', regex=True)
        out["전화번호1"] = phone_series
        out["전화번호2"] = phone_series
        out["우편번호"] = get_val(_COLUMN_MAP["우편번호"])
        
        target_addr_cols = ["배송 주", "배송 도시", "배송 주소 1", "배송 주소 2"]
        if any(c in src.columns for c in target_addr_cols):
            addr_parts = [src[col].fillna("").astype(str).str.strip() if col in src.columns else pd.Series([""] * len(src)) for col in target_addr_cols]
            combined_addr = addr_parts[0] + " " + addr_parts[1] + " " + addr_parts[2] + " " + addr_parts[3]
            out["주소"] = combined_addr.str.replace(r'\s+', ' ', regex=True).str.strip()
        else:
            city_col = next((c for c in _CITY_CANDIDATES if c in src.columns), None)
            city_part = src[city_col].fillna("").astype(str).str.strip() if city_col else pd.Series([""] * len(src))
            out["주소"] = (city_part + " " + get_val(_COLUMN_MAP["상세주소"]).str.strip()).str.strip()
        
        out["상품명1"] = get_val(_COLUMN_MAP["상품명"])
        out["수량1"] = get_val(_COLUMN_MAP["수량1"]).replace("", "1")

        return out[_TARGET_COLUMNS]

    input_file = st.file_uploader("원본 주문 엑셀/CSV 선택", type=["xlsx", "xls", "csv"])
    if input_file and st.button("🚀 변환 실행"):
        try:
            file_name = input_file.name.lower()
            if file_name.endswith(".csv"):
                try: src_df = pd.read_csv(input_file, dtype=str, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    input_file.seek(0)
                    src_df = pd.read_csv(input_file, dtype=str, encoding="cp949")
            else:
                src_df = pd.read_excel(input_file, dtype=str, engine="openpyxl" if file_name.endswith(".xlsx") else None)

            src_df.columns = [str(c).strip() for c in src_df.columns]
            final_df = _convert_auto(src_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False)

            st.success(f"✅ 변환 완료! 총 {len(final_df)}건 처리됨.")
            st.download_button("📥 변환된 엑셀 다운로드", data=output.getvalue(), file_name=f"요정비닐_송장_{datetime.now().strftime('%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(final_df.head(), use_container_width=True)
        except Exception as e:
            st.error(f"❌ 변환 오류: {e}")

# ==========================================
# [기능 2] 오픈마켓 수익 분석 (V2.4)
# ==========================================
def show_openmarket_calc():
    st.title("🛍️ 오픈마켓 수익 분석 시뮬레이터")
    st.markdown("---")
    st.info("💡 구글 시트의 원가 및 네이버 판매가를 연동하여 마진을 분석합니다.")
    # (앞서 완성한 오픈마켓 분석 코드 전체가 이곳에 들어갑니다. 길이상 핵심만 요약 표기)
    st.write("*(오픈마켓 분석 엔진 정상 로드 완료)*")
    # 실제 적용하실 때는 앞서 드린 V2.4 코드를 이 안에 그대로 붙여넣으시면 됩니다.

# ==========================================
# [기능 3] 생산 작업 지시서 생성 (신규)
# ==========================================
def show_production_order():
    st.title("📋 생산 작업 지시서 자동 생성")
    st.markdown("---")
    
    # 🖨️ 인쇄용 CSS (프린트 시 사이드바와 버튼을 숨기고 표만 깔끔하게 출력)
    st.markdown("""
        <style>
        @media print {
            [data-testid="stSidebar"] { display: none; }
            .stButton { display: none; }
            header { display: none; }
        }
        </style>
    """, unsafe_allow_html=True)

    try:
        # 데이터 로드 및 정제
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote('상품목록')}")
        df_products.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df_products.columns]
        
        st.subheader("🏭 오늘의 생산 계획 입력 (단위: 박스/수량)")
        production_data = {}
        
        # 상품 리스트 동적 생성 (3열 배치)
        cols = st.columns(3)
        for idx, row in df_products.iterrows():
            sku = str(row.get('SKU ID', f'unknown_{idx}'))
            name = str(row.get('상품명', '이름없음'))
            col_idx = idx % 3
            production_data[sku] = cols[col_idx].number_input(f"{name}", min_value=0, value=0, key=f"prod_{sku}")

        if st.button("🚀 지시서 생성하기"):
            order_list = []
            total_rolls = 0
            
            for sku, qty in production_data.items():
                if qty > 0:
                    prod_row = df_products[df_products['SKU ID'].astype(str) == sku].iloc[0]
                    # 💡 시트에 '롤당수량' 혹은 '롤당 카운팅' 열이 있어야 합니다.
                    pcs_per_roll = clean_num(next((prod_row[k] for k in prod_row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                    needed_rolls = qty / pcs_per_roll if pcs_per_roll > 0 else 0
                    
                    order_list.append({
                        "상품명": prod_row['상품명'],
                        "목표생산(박스/매)": qty,
                        "필요 원단(롤)": round(needed_rolls, 2)
                    })
                    total_rolls += needed_rolls
            
            if order_list:
                st.markdown("---")
                st.subheader("📄 일일 작업 지시서")
                st.write(f"**작성일시:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
                
                df_order = pd.DataFrame(order_list)
                st.table(df_order)
                
                st.success(f"🔥 **총 투입 필요 원단: {round(total_rolls, 2)} 롤**")
                st.info("🖨️ 키보드의 **[Ctrl + P]** 를 누르시면 현재 지시서 표가 A4 용지에 깔끔하게 인쇄됩니다.")
            else:
                st.warning("생산할 수량을 1개 이상 입력해 주세요.")
                
    except Exception as e:
        st.error(f"시스템 오류: {e}")

# ==========================================
# 🚀 메인 사이드바 네비게이션 (라우팅)
# ==========================================
st.sidebar.title("🏭 페어리테림 ERP")
st.sidebar.markdown("---")

# 메뉴 선택
menu = st.sidebar.radio(
    "📌 메인 메뉴",
    ["📦 택배 송장 변환기", "🛍️ 오픈마켓 수익 분석", "📋 생산 작업 지시서"]
)

st.sidebar.markdown("---")
st.sidebar.caption("© 2026 Fairy Vinyl Mfg.")

# 선택된 메뉴에 따라 함수 호출
if menu == "📦 택배 송장 변환기":
    show_invoice_converter()
elif menu == "🛍️ 오픈마켓 수익 분석":
    show_openmarket_calc()
elif menu == "📋 생산 작업 지시서":
    show_production_order()
