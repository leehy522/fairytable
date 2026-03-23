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

        # 4. 마진 계산 함수
        def calc_row(row):
            try:
                # 콤마(,)나 문자가 섞여 있어도 숫자로 강제 변환하는 함수
                def clean_num(value):
                    if pd.isna(value): return 0
                    # 문자열로 바꾼 뒤 콤마 제거하고 숫자만 남김
                    s = str(value).replace(',', '').strip()
                    return pd.to_numeric(s, errors='coerce')
                    
                # 데이터 타입 강제 변환
                garo = pd.to_numeric(row['가로(cm)'], errors='coerce')
                sero = pd.to_numeric(row['세로(cm)'], errors='coerce')
                dukki = pd.to_numeric(row['두께(T)'], errors='coerce')
                maesu = pd.to_numeric(row['매수'], errors='coerce')
                s_ratio = pd.to_numeric(row['신재비율'], errors='coerce')
                j_ratio = pd.to_numeric(row['안료비율'], errors='coerce')

                # 1장 원가 및 수익 계산
                one_cost = (s_ratio * sinjae + j_ratio * jaesaeng + 
                            (im_sinjae if s_ratio > 0 else im_jaesaeng)) * \
                           (garo * sero * dukki * 0.00000184)
                
                total_cost = round(one_cost * maesu, 0)
                nap_ga = pd.to_numeric(row['쿠팡 로켓 납품가(부가세 별도)'], errors='coerce')
                pan_ga = pd.to_numeric(row['쿠팡 판매가'], errors='coerce')
                profit = nap_ga - total_cost
                
                return pd.Series([total_cost, nap_ga, pan_ga, profit])
            except:
                return pd.Series([0, 0, 0, 0])

        # 5. 결과 적용 및 출력 (요청하신 4개 항목 위주)
        result_cols = ['원가(1장*매수)', '쿠팡 로켓 납품가(부가세 별도)', '쿠팡 판매가', '수익']
        # 계산 결과를 새 컬럼으로 추가
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
