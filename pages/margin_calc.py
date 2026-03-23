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
        anlyo_price = float(target_cost['안료'])

        # 4. 마진 계산 내부 함수 (들여쓰기 수정됨)
        def calc_row(row):
            try:
                def clean_num(value):
                    if pd.isna(value): return 0
                    s = str(value).replace(',', '').replace('%', '').strip()
                    return pd.to_numeric(s, errors='coerce')

                # 1. 데이터 추출
                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                
                # 롤 전체에서 나오는 총 '장'수 또는 '세트'수
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and '수량' in k), 1))
                box_cost = clean_num(row.get('박스비', 0))

                # 2. 1kg당 원재료 단가 (선택된 월 기준)
                s_ratio = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 0))
                j_ratio = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_ratio = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                
                if s_ratio > 1: s_ratio /= 100
                if j_ratio > 1: j_ratio /= 100
                if a_ratio > 1: a_ratio /= 100

                material_unit_price = (sinjae * s_ratio) + (jaesaeng * j_ratio) + (anlyo_price * a_ratio)

                # 3. [교정된 윤겸님 공식]
                # ① 롤 전체 무게 (kg)
                total_roll_weight = garo * sero * dukki * 0.00000184 * length
                
                # ② 롤 전체 가격 (원)
                total_roll_price = total_roll_weight * material_unit_price
                
                # ③ 최종 상품원가 (한 박스 기준)
                # 롤 전체 가격을 롤당 수량으로 나누면 '한 세트(박스)'의 원단값이 나옵니다.
                divisor = total_pcs_in_roll if total_pcs_in_roll > 0 else 1
                total_cost = round((total_roll_price / divisor) + box_cost, 0)
                
                # 4. 수익 계산
                nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                pan_ga = clean_num(next((row[k] for k in row.index if '판매가' in k and '쿠팡' in k), 0))
                profit = nap_ga - total_cost
                
                return pd.Series([total_cost, nap_ga, pan_ga, profit])
            except:
                return pd.Series([0, 0, 0, 0])
                
        # 5. 결과 적용 및 필터링 출력
        result_cols = ['최종 상품원가', '쿠팡 로켓 납품가(부가세 별도)', '쿠팡 판매가', '수익']
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
