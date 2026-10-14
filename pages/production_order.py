import streamlit as st
import pandas as pd
import re
import unicodedata
import math
import pdfplumber
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
                    
                    headers = [str(cell).replace('\n', '').strip() for cell in table[0] if cell]
                    sku_idx, qty_idx = -1, -1
                    
                    for i, h in enumerate(headers):
                        h_upper = h.upper()
                        if 'SKU' in h_upper or '상품코드' in h_upper or '상품 ID' in h_upper:
                            sku_idx = i
                        if '수량' in h_upper or 'QTY' in h_upper or '주문수량' in h_upper:
                            qty_idx = i
                    
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

def parse_excel_order(file, ext):
    """엑셀/CSV 발주서에서 데이터를 추출하여 SKU별 수량을 합산하는 함수"""
    sku_qty_map = {}
    try:
        # 파일 확장자에 따른 로드
        if ext == "csv":
            try:
                df = pd.read_csv(file, dtype=str, encoding="utf-8-sig")
            except UnicodeDecodeError:
                file.seek(0)
                df = pd.read_csv(file, dtype=str, encoding="cp949")
        else:
            df = pd.read_excel(file, dtype=str)
        
        # 헤더(컬럼명)에서 SKU와 수량 열을 유연하게 탐색
        sku_col, qty_col = None, None
        for col in df.columns:
            c_str = str(col).replace('\n', '').strip().upper()
            
            # 제공해주신 엑셀 파일의 '상품바코드', '상품번호' 등을 우선 스캔
            if not sku_col and any(x in c_str for x in ['바코드', '상품코드', 'SKU', '상품 ID', '상품번호']):
                sku_col = col
            # '발주수량', '확정수량' 등을 스캔
            if not qty_col and any(x in c_str for x in ['발주수량', '주문수량', '확정수량', '수량', 'QTY']):
                qty_col = col
                
        if sku_col and qty_col:
            for _, row in df.iterrows():
                sku_val = str(row[sku_col]).strip()
                sku = re.sub(r'\.0$', '', sku_val).strip()
                qty_val = clean_num(row[qty_col])
                
                if sku and qty_val > 0:
                    sku_qty_map[sku] = sku_qty_map.get(sku, 0) + int(qty_val)
        else:
            st.warning("⚠️ 엑셀 파일에서 '상품바코드/번호' 또는 '발주수량' 열을 찾을 수 없습니다. 양식을 확인해 주세요.")
            
        return sku_qty_map
    except Exception as e:
        st.error(f"엑셀 분석 중 오류가 발생했습니다: {e}")
        return {}

def show_production_order():
    # 💡 화면 전환 및 데이터 유지를 위한 상태 변수 초기화
    if "prod_step" not in st.session_state:
        st.session_state.prod_step = 1
    if "saved_prod_data" not in st.session_state:
        st.session_state.saved_prod_data = {}

    # 🖨️ 인쇄용 CSS (사이드바, 버튼, 최상단 헤더 모두 숨김)
    st.markdown("""
        <style>
        @media print {
            [data-testid="stSidebar"] { display: none !important; }
            .stButton { display: none !important; }
            .stFileUploader { display: none !important; }
            header { display: none !important; }
            footer { display: none !important; }
            h1 { display: none !important; } 
            .stAlert { display: none !important; }
        }
        </style>
    """, unsafe_allow_html=True)

    try:
        # 데이터 로드
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name = quote("상품목록")
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}")
        df_products.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df_products.columns]
        
        # -----------------------------------------------------------
        # [STEP 1] 데이터 입력 화면 
        # -----------------------------------------------------------
        if st.session_state.prod_step == 1:
            st.title("📋 생산 작업 지시서 자동 생성")
            st.markdown("---")
            
            st.subheader("📑 발주서 자동 인식 (PDF 및 엑셀)")
            st.info("발주서(PDF, 엑셀, CSV)를 올리면 상품코드(SKU/바코드)와 수량을 읽어와 아래 입력창에 자동으로 채워줍니다.")
            
            # 💡 PDF 뿐만 아니라 xlsx, xls, csv 파일도 업로드 가능하게 수정
            uploaded_file = st.file_uploader("발주서 파일 업로드", type=["pdf", "xlsx", "xls", "csv"])
            
            if uploaded_file and st.button("🔄 데이터 불러오기"):
                extracted_data = {}
                file_ext = uploaded_file.name.split('.')[-1].lower()
                
                # 확장자에 따라 분석 엔진 분기 처리
                if file_ext == "pdf":
                    extracted_data = parse_pdf_order(uploaded_file)
                elif file_ext in ["xlsx", "xls", "csv"]:
                    extracted_data = parse_excel_order(uploaded_file, file_ext)
                    
                if extracted_data:
                    for sku, qty in extracted_data.items():
                        current = st.session_state.saved_prod_data.get(sku, 0)
                        st.session_state.saved_prod_data[sku] = current + qty
                    st.success(f"✅ 총 {len(extracted_data)}종의 상품 수량이 성공적으로 입력되었습니다!")
                    st.rerun()  # 화면 새로고침
                elif file_ext == "pdf":
                    st.warning("⚠️ 표 데이터를 찾지 못했습니다.")

            st.markdown("---")
            st.subheader("🏭 오늘의 생산 계획 입력 (단위: 수량)")
            
            cols = st.columns(3)
            all_skus = []
            
            for idx, row in df_products.iterrows():
                sku = str(row.get('SKU ID', f'unknown_{idx}'))
                sku = re.sub(r'\.0$', '', sku).strip() 
                name = str(row.get('상품명', '이름없음'))
                all_skus.append(sku)
                
                col_idx = idx % 3
                current_val = st.session_state.saved_prod_data.get(sku, 0)
                
                cols[col_idx].number_input(
                    f"{name}", 
                    min_value=0, 
                    value=current_val,
                    key=f"ui_prod_{sku}"
                )

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚀 지시서 생성 및 인쇄 화면으로 이동", type="primary"):
                order_list = []
                total_rolls = 0
                
                for sku in all_skus:
                    val = st.session_state[f"ui_prod_{sku}"]
                    st.session_state.saved_prod_data[sku] = val
                    
                    if val > 0:
                        prod_row = df_products[df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() == sku].iloc[0]
                        pcs_per_roll_col = next((k for k in prod_row.index if '상품' in k and '개수' in k and '1롤' in k), None)
                        pcs_per_roll = clean_num(prod_row[pcs_per_roll_col]) if pcs_per_roll_col else 1
                        
                        raw_needed_rolls = val / pcs_per_roll if pcs_per_roll > 0 else 0
                        needed_rolls = int(math.ceil(raw_needed_rolls)) 
                        
                        order_list.append({
                            "상품명": prod_row.get('상품명', '알 수 없음'),
                            "목표 생산(수량)": f"{val:,}", 
                            "필요 원단(롤)": f"{needed_rolls} 롤"
                        })
                        total_rolls += needed_rolls
                
                if order_list:
                    st.session_state.order_list = order_list
                    st.session_state.total_rolls = total_rolls
                    st.session_state.prod_step = 2 
                    st.rerun()
                else:
                    st.warning("생산할 수량을 1개 이상 입력해 주세요.")

        # -----------------------------------------------------------
        # [STEP 2] 인쇄 전용 화면
        # -----------------------------------------------------------
        elif st.session_state.prod_step == 2:
            st.subheader("📄 일일 생산 작업 지시서")
            st.write(f"**작성일시:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            
            df_order = pd.DataFrame(st.session_state.order_list)
            st.table(df_order)
            
            st.success(f"🔥 **총 투입 필요 원단: {st.session_state.total_rolls} 롤**")
            st.info("🖨️ 키보드의 **[Ctrl + P]** 를 누르시면 현재 화면의 표만 A4 용지에 깔끔하게 인쇄됩니다.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            
            if col1.button("🔄 내용 수정하기"):
                st.session_state.prod_step = 1
                st.rerun()
            
            if col2.button("🗑️ 모두 지우고 새로 작성"):
                st.session_state.saved_prod_data = {}
                st.session_state.prod_step = 1
                st.rerun()

    except Exception as e:
        st.error(f"시스템 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    if check_password():
        show_production_order()
