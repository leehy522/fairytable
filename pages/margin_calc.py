import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password

def show_margin_calc():
    st.title("💰 월별 마진 시뮬레이션")

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        
        # 1. 한글 시트 이름을 URL용으로 변환
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        url_products = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}"
        url_costs = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}"
        
        # 2. 데이터 로드 및 전처리
        df_products = pd.read_csv(url_products)
        df_costs = pd.read_csv(url_costs)
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        # 3. 입력값 설정
        months = df_costs['월'].unique().tolist()
        selected_month = st.selectbox("분석할 월을 선택하세요", months)

        # 해당 월의 원가 행 추출
        target_cost = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        sinjae = float(target_cost['신재'])
        jaesaeng = float(target_cost['재생'])
        im_sinjae = float(target_cost['임가공(신재)'])
        im_jaesaeng = float(target_cost['임가공(재생)'])
        anlyo_price = float(target_cost['안료'])

        # 4. 마진 계산 내부 함수 (들여쓰기 수정됨)
        def calc_row(row):
            try:
                def clean_num(value):
                    if pd.isna(value): return 0
                    s = str(value).replace(',', '').strip()
                    return pd.to_numeric(s, errors='coerce')

                # 데이터 세척
                garo = clean_num(row['가로'])
                sero = clean_num(row['세로'])
                dukki = clean_num(row['두께'])
                maesu = clean_num(row['매수'])
                s_ratio = clean_num(row['신재비율'])
                j_ratio = clean_num(row['재생비율'])
                a_ratio = clean_num(row['안료비율'])

                # 원료비 + 가공비 공식 적용
                material_cost = (s_ratio * sinjae) + (j_ratio * jaesaeng) + (a_ratio * anlyo_price)
                processing_fee = im_sinjae if s_ratio > 0 else im_jaesaeng

                # 1장 원가 = (원료비 + 가공비) * (부피 * 비중)
                one_cost = (material_cost + processing_fee) * (garo * sero * dukki * 0.00000184)
                
                total_cost = round(one_cost * maesu, 0)
                nap_ga = clean_num(row['쿠팡 로켓 납품가(부가세 별도)'])
                pan_ga = clean_num(row['쿠팡 판매가'])
                profit = nap_ga - total_cost
                
                return pd.Series([total_cost, nap_ga, pan_ga, profit])
            except:
                return pd.Series([0, 0, 0, 0])

        # 5. 결과 적용 및 필터링 출력
        result_cols = ['원가(1장*매수)', '쿠팡 로켓 납품가(부가세 별도)', '쿠팡 판매가', '수익']
        df_res = df_products.apply(calc_row, axis=1)
        df_products[result_cols] = df_res

        # 최종 화면 출력
        display_cols = ['상품명'] + result_cols
        st.subheader(f"📊 {selected_month} 마진 분석 결과")
        st.dataframe(df_products[display_cols], use_container_width=True)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")

# 인증 확인 후 실행
if check_password():
    show_margin_calc()
