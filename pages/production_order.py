import streamlit as st
import pandas as pd
import re
import unicodedata
import math  # 💡 올림 처리를 위한 수학 라이브러리 추가
from datetime import datetime
from urllib.parse import quote
from auth import check_password

def clean_num(value):
    """문자열에서 숫자만 추출하여 float 형태로 반환하는 헬퍼 함수"""
    if pd.isna(value) or value == '': 
        return 0
    s = re.sub(r'[^0-9.]', '', str(value))
    return pd.to_numeric(s, errors='coerce') if s else 0

def show_production_order():
    st.title("📋 생산 작업 지시서 자동 생성")
    st.markdown("---")
    
    # 🖨️ 인쇄용 CSS (프린트 시 사이드바와 버튼 등 불필요한 요소를 숨기고 표만 깔끔하게 출력)
    st.markdown("""
        <style>
        @media print {
            [data-testid="stSidebar"] { display: none !important; }
            .stButton { display: none !important; }
            header { display: none !important; }
            footer { display: none !important; }
        }
        </style>
    """, unsafe_allow_html=True)

    try:
        # 1. 데이터 로드 및 정제
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name = quote("상품목록")
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}")
        
        # 맥(Mac) 등에서 발생하는 한글 자소 분리 현상 방지 및 공백 제거
        df_products.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df_products.columns]
        
        st.subheader("🏭 오늘의 생산 계획 입력 (단위: 수량)")
        production_data = {}
        
        # 2. 상품 리스트 동적 입력 폼 생성 (3열 배치)
        cols = st.columns(3)
        for idx, row in df_products.iterrows():
            sku = str(row.get('SKU ID', f'unknown_{idx}'))
            sku = re.sub(r'\.0$', '', sku).strip() 
            name = str(row.get('상품명', '이름없음'))
            
            col_idx = idx % 3
            production_data[sku] = cols[col_idx].number_input(
                f"{name}", 
                min_value=0, 
                value=0, 
                key=f"prod_{sku}_{idx}"
            )

        # 3. 작업 지시서 계산 및 출력 로직
        if st.button("🚀 지시서 생성하기"):
            order_list = []
            total_rolls = 0
            
            for sku, qty in production_data.items():
                if qty > 0:
                    prod_row = df_products[df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() == sku].iloc[0]
                    
                    # 시트의 "상품 개수/1롤" 열 탐색 및 연동
                    pcs_per_roll_col = next((k for k in prod_row.index if '상품' in k and '개수' in k and '1롤' in k), None)
                    
                    if pcs_per_roll_col:
                        pcs_per_roll = clean_num(prod_row[pcs_per_roll_col])
                    else:
                        pcs_per_roll = 1
                        st.warning(f"⚠️ '{prod_row.get('상품명')}'의 '상품 개수/1롤' 데이터를 시트에서 찾을 수 없어 임시로 1로 계산했습니다.")
                    
                    # 💡 필요 원단 롤 수 계산 (무조건 올림 처리 후 정수 변환)
                    raw_needed_rolls = qty / pcs_per_roll if pcs_per_roll > 0 else 0
                    needed_rolls = int(math.ceil(raw_needed_rolls)) 
                    
                    order_list.append({
                        "상품명": prod_row.get('상품명', '알 수 없음'),
                        "목표 생산(수량)": f"{qty:,}",  # 천 단위 콤마 추가로 가독성 향상
                        "필요 원단(롤)": f"{needed_rolls} 롤"  # 소수점 제거 및 '롤' 텍스트 고정
                    })
                    total_rolls += needed_rolls
            
            # 결과 표 출력
            if order_list:
                st.markdown("---")
                st.subheader("📄 일일 생산 작업 지시서")
                st.write(f"**작성일시:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
                
                df_order = pd.DataFrame(order_list)
                st.table(df_order)
                
                # 총 투입 필요 원단도 정수(int)로 깔끔하게 출력
                st.success(f"🔥 **총 투입 필요 원단: {total_rolls} 롤**")
                st.info("🖨️ 키보드의 **[Ctrl + P]** 를 누르시면 현재 화면의 표와 수량만 A4 용지에 깔끔하게 인쇄됩니다.")
            else:
                st.warning("생산할 수량을 1개 이상 입력해 주세요.")
                
    except Exception as e:
        st.error(f"시스템 오류가 발생했습니다: {e}")

# 단독으로 파일 실행 시 테스트를 위한 로직
if __name__ == "__main__":
    if check_password():
        show_production_order()
