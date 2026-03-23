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
            # 1장 원가 (비중 반영)
            one_cost = (row['신재비율'] * sinjae + row['재생비율'] * jaesaeng + 
                        (im_sinjae if row['신재비율'] > 0 else im_jaesaeng)) * \
                       (row['가로'] * row['세로'] * row['두께'] * 0.00000184)
            
            total_cost = round(one_cost * row['매수'], 0)
            nap_ga = row['쿠팡 로켓 납품가(부가세 별도)']
            pan_ga = row['쿠팡 판매가']
            profit = nap_ga - total_cost
            
            return pd.Series([total_cost, nap_ga, pan_ga, profit])

        # 결과 계산 적용
        result_cols = ['원가(1장*매수)', '쿠팡 로켓 납품가(부가세 별도)', '쿠팡 판매가', '수익']
        df_products[result_cols] = df_products.apply(calc_row, axis=1)

        # 5. 최종 출력 (요청하신 4개 항목 + 상품명)
        display_df = df_products[['상품명'] + result_cols]
        st.dataframe(display_df, use_container_width=True)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다.")
        st.info(f"상세 에러: {e}")

if check_password():
    show_margin_calc()
