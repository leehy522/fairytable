import streamlit as st
import pandas as pd
import re
import unicodedata
import math
import io
import pdfplumber  # 💡 PDF 텍스트 및 표 추출 라이브러리 추가
from datetime import datetime
from urllib.parse import quote
from auth import check_password

def clean_num(value):
    """문자열에서 숫자만 추출하여 float 형태로 반환하는 헬퍼 함수"""
    if pd.isna(value) or value == '': 
        return 0
    s = re.sub(r'[^0-9.]', '', str(value))
    return pd.to_numeric(s, errors='coerce') if s else 0

def parse_pdf_order(file):
    """PDF 발주서에서 표 데이터를 추출하여 SKU별 수량을 합산하는 함수"""
    sku_qty_map = {}
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table: continue
                    
                    # 헤더(첫 번째 행)에서 SKU와 수량 열의 인덱스 찾기
                    headers = [str(cell).replace('\n', '').strip() for cell in table[0] if cell]
                    sku_idx, qty_idx = -1, -1
                    
                    for i, h in enumerate(headers):
                        h_upper = h.upper()
                        # 발주서의 컬럼명 포맷에 맞춰 유연하게 키워드 지정
                        if 'SKU' in h_upper or '상품코드' in h_upper or '상품 ID' in h_upper:
                            sku_idx = i
                        if '수량' in h_upper or 'QTY' in h_upper or '주문수량' in h_upper:
                            qty_idx = i
                    
                    # SKU와 수량 컬럼을 모두 찾았다면 데이터 추출
                    if sku_idx != -1 and qty_idx != -1:
                        for row in table[1:]:
                            if len(row) > max(sku_idx, qty_idx):
                                sku_val = str(row[sku_idx]).strip()
                                sku = re.sub(r'\.0$', '', sku_val).strip()
                                
                                qty_val = clean_num(row[qty_idx])
                                
                                if sku and qty_val > 0:
                                    sku_qty_map[sku] = sku_qty_map.get(sku, 0) + int(qty_val)
        return sku_qty_map
    except Exception as e:
        st.error(f"PDF 분석 중 오류가 발생했습니다: {e}")
        return {}

def show_production_order():
    st.title("📋 생산 작업 지시서 자동 생성")
    st.markdown("---")
    
    # 🖨️ 인쇄용 CSS (프린트 시 불필요한 요소를 숨기고 표만 출력)
    st.markdown("""
        <style>
        @media print {
            [data-testid="stSidebar"] { display: none !important; }
            .stButton { display: none !important; }
            .stFileUploader { display: none !important; }
            header { display: none !important; }
            footer { display: none !important; }
            .stAlert { display: none !important; }
        }
        </style>
    """, unsafe_allow_html=True)

    try:
        # 1. 데이터 로드 및 정제
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name = quote("상품목록")
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}")
        df_products.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df_products.columns]
        
        # 💡 PDF 업로드 UI 추가
        st.subheader("📑 PDF 발주서 자동 인식 (선택 사항)")
        st.info("발주서 PDF를 올리면 상품코드(SKU)와 수량을 읽어와 아래 입력창에 자동으로 채워줍니다.")
        
        uploaded_pdf = st.file_uploader("발주서 PDF 파일 업로드", type=["pdf"])
        if uploaded_pdf:
            if st.button("🔄 PDF 데이터 불러오기"):
                extracted_data = parse_pdf_order(uploaded_pdf)
                if extracted_data:
                    # 추출된 데이터를 Streamlit Session State에 업데이트
                    for sku, qty in extracted_data.items():
                        st.session_state[f"prod_{sku}"] = st.session_state.get(f"prod_{sku}", 0) + qty
                    st.success(f"✅ 총 {len(extracted_data)}종의 상품 수량이 자동 입력되었습니다! 아래 수량을 확인 후 지시서를 생성하세요.")
                else:
                    st.warning("⚠️ PDF에서 'SKU' 및 '수량' 표 데이터를 찾지 못했습니다. 파일 양식을 확인해 주세요.")
        
        st.markdown("---")
        st.subheader("🏭 오늘의 생산 계획 입력 (단위: 수량)")
        
        # 2. 상품 리스트 동적 입력 폼 생성
        production_data = {}
        cols = st.columns(3)
        for idx, row in df_products.iterrows():
            sku = str(row.get('SKU ID', f'unknown_{idx}'))
            sku = re.sub(r'\.0$', '', sku).strip() 
            name = str(row.get('상품명', '이름없음'))
            
            # Session State 키를 고유하게 설정하여 PDF 데이터와 연동
            widget_key = f"prod_{sku}"
            if widget_key not in st.session_state:
                st.session_state[widget_key] = 0
                
            col_idx = idx % 3
            production_data[sku] = cols[col_idx].number_input(
                f"{name}", 
                min_value=0, 
                key=widget_key
            )

        # 3. 작업 지시서 계산 및 출력 로직
        if st.button("🚀 지시서 생성하기"):
            order_list = []
            total_rolls = 0
            
            for sku, qty in production_data.items():
                if qty > 0:
                    prod_row = df_products[df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() == sku].iloc[0]
                    
                    pcs_per_roll_col = next((k for k in prod_row.index if '상품' in k and '개수' in k and '1롤' in k), None)
                    if pcs_per_roll_col:
                        pcs_per_roll = clean_num(prod_row[pcs_per_roll_col])
                    else:
                        pcs_per_roll = 1
                        st.warning(f"⚠️ '{prod_row.get('상품명')}'의 '상품 개수/1롤' 데이터를 찾을 수 없어 1로 계산했습니다.")
                    
                    raw_needed_rolls = qty / pcs_per_roll if pcs_per_roll > 0 else 0
                    needed_rolls = int(math.ceil(raw_needed_rolls)) 
                    
                    order_list.append({
                        "상품명": prod_row.get('상품명', '알 수 없음'),
                        "목표 생산(수량)": f"{qty:,}", 
                        "필요 원단(롤)": f"{needed_rolls} 롤"
                    })
                    total_rolls += needed_rolls
            
            # 결과 표 출력
            if order_list:
                st.markdown("---")
                st.subheader("📄 일일 생산 작업 지시서")
                st.write(f"**작성일시:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
                
                df_order = pd.DataFrame(order_list)
                st.table(df_order)
                
                st.success(f"🔥 **총 투입 필요 원단: {total_rolls} 롤**")
                st.info("🖨️ 키보드의 **[Ctrl + P]** 를 누르시면 현재 화면의 표와 수량만 A4 용지에 깔끔하게 인쇄됩니다.")
            else:
                st.warning("생산할 수량을 1개 이상 입력해 주세요.")
                
    except Exception as e:
        st.error(f"시스템 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    if check_password():
        show_production_order()
