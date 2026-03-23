import streamlit as st
import pandas as pd
from urllib.parse import quote # 한글 주소 변환을 위해 추가
from auth import check_password

def show_margin_calc():
    st.title("💰 월별 마진 시뮬레이션")

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        
        # 한글 시트 이름을 URL용으로 변환
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        url_products = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}"
        url_costs = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}"
        
        # 이후 로직은 동일
        df_products = pd.read_csv(url_products)
        df_costs = pd.read_csv(url_costs)

        # 3. 입력값 설정
        months = df_costs['월'].unique().tolist()
        selected_month = st.selectbox("분석할 월을 선택하세요", months)

        # 해당 월의 원가 행 추출
        target_cost = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # 원가 요소 할당
        sinjae = float(target_cost['신재'])
        jaesaeng = float(target_cost['재생'])
        im_sinjae = float(target_cost['임가공(신재)'])
        im_jaesaeng = float(target_cost['임가공(재생)'])

# 4. 마진 계산 함수
        def calc_row(row):
            try:
                # 시트에서 읽어온 값들을 숫자로 강제 변환 (오류 방지 핵심)
                garo = pd.to_numeric(row['가로'], errors='coerce')
                sero = pd.to_numeric(row['세로'], errors='coerce')
                dukki = pd.to_numeric(row['두께'], errors='coerce')
                maesu = pd.to_numeric(row['매수'], errors='coerce')
                
                # 비율 및 단가 (이미 float 변환됨)
                s_ratio = pd.to_numeric(row['신재비율'], errors='coerce')
                j_ratio = pd.to_numeric(row['재생비율'], errors='coerce')

                # 1장 원가 계산 (숫자 데이터로만 연산 수행)
                one_cost = (s_ratio * sinjae + j_ratio * jaesaeng + 
                            (im_sinjae if s_ratio > 0 else im_jaesaeng)) * \
                           (garo * sero * dukki * 0.00000184)
                
                total_cost = round(one_cost * maesu, 0)
                nap_ga = pd.to_numeric(row['쿠팡 로켓 납품가(부가세 별도)'], errors='coerce')
                pan_ga = pd.to_numeric(row['쿠팡 판매가'], errors='coerce')
                profit = nap_ga - total_cost
                
                return pd.Series([total_cost, nap_ga, pan_ga, profit])
            except Exception as e:
                # 계산 중 오류 발생 시 0으로 반환하여 시스템 멈춤 방지
                return pd.Series([0, 0, 0, 0])

        # 결과 계산 적용
        result_cols = ['원가(1장*매수)', '쿠팡 로켓 납품가(부가세 별도)', '쿠팡 판매가', '수익']
        df_products[result_cols] = df_products.apply(calc_row, axis=1)
    show_margin_calc()
