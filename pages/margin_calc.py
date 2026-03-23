import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from auth import check_password

# 페이지 설정
st.set_page_config(page_title="마진 시뮬레이션", layout="wide")

def show_margin_calc():
    st.title("💰 월별 마진 시뮬레이션")
    
    try:
        # 1. 데이터 불러오기 (변수 정의 단계)
        conn = st.connection("gsheets", type=GSheetsConnection)
        
        # 시트 이름은 실제 귀하의 구글 시트 탭 이름과 일치해야 합니다.
        df_products = conn.read(worksheet="상품목록") 
        df_costs = conn.read(worksheet="원가기준")
        
        if df_products.empty or df_costs.empty:
            st.error("시트에서 데이터를 불러오지 못했습니다.")
            return

        # 2. 입력값 설정
        months = df_costs['월'].unique().tolist()
        selected_month = st.selectbox("분석할 월을 선택하세요", months)

        # 3. 계산 및 필터링 로직
        # 컬럼명 공백 제거
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()

        target_cost = df_costs[df_costs['월'] == selected_month]
        
        # 원가 요소 할당
        sinjae = float(target_cost['신재'].values[0])
        jaesaeng = float(target_cost['재생'].values[0])
        im_sinjae = float(target_cost['임가공(신재)'].values[0])
        im_jaesaeng = float(target_cost['임가공(재생)'].values[0])

        # 1장 원가 및 요청 항목 계산
        def calc_row(row):
            one_cost = (row['신재비율'] * sinjae + row['재생비율'] * jaesaeng + 
                        (im_sinjae if row['신재비율'] > 0 else im_jaesaeng)) * \
                       (row['가로'] * row['세로'] * row['두께'] * 0.00000184)
            
            total_cost = round(one_cost * row['매수'], 0)
            # 시트 컬럼명과 정확히 일치해야 함
            nap_ga = row['쿠팡 로켓 납품가(부가세 별도)']
            pan_ga = row['쿠팡 판매가']
            profit = nap_ga - total_cost
            
            return pd.Series([total_cost, nap_ga, pan_ga, profit])

        # 결과 데이터프레임 생성
        result_cols = ['원가(1장*매수)', '쿠팡 로켓 납품가(부가세 별도)', '쿠팡 판매가', '수익']
        df_products[result_cols] = df_products.apply(calc_row, axis=1)

        # 4. 최종 출력 (요청하신 항목만)
        display_df = df_products[['상품명'] + result_cols]
        st.dataframe(display_df, use_container_width=True)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")

# 인증 후 실행
if check_password():
    show_margin_calc()
