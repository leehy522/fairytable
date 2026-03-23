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
                    s = str(value).replace(',', '').replace('%', '').strip()
                    return pd.to_numeric(s, errors='coerce')

                # 1. 제품 규격 및 매수 데이터 세척
                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                maesu = clean_num(row.get('매수', 0))
                
                # 2. 원재료 배합 비율 추출
                s_ratio = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 0))
                j_ratio = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_ratio = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                
                # 비율 정규화 (2.7 -> 0.027)
                if s_ratio > 1: s_ratio /= 100
                if j_ratio > 1: j_ratio /= 100
                if a_ratio > 1: a_ratio /= 100

                # 3. 순수 원재료비 공식 (임가공비 제외)
                # 1kg당 원재료비 = (신재비율*신재단가) + (재생비율*재생단가) + (안료비율*안료단가)
                material_price_per_kg = (s_ratio * sinjae) + (j_ratio * jaesaeng) + (a_ratio * anlyo_price)
                
                # 4. 제품 1장당 중량 및 최종 원가 계산
                # 1장당 중량(kg) = 가로 * 세로 * 두께 * 비중(0.00000184)
                weight_per_piece = garo * sero * dukki * 0.00000184
                
                # 최종 원가 = 1kg당 원재료비 * 1장당 중량 * 매수
                total_cost = round(material_price_per_kg * weight_per_piece * maesu, 0)
                
                # 5. 수익 계산
                nap_ga = clean_num(row.get('쿠팡 로켓 납품가(부가세 별도)', 0))
                pan_ga = clean_num(row.get('쿠팡 판매가', 0))
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
