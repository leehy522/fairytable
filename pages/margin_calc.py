import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io # 엑셀 변환을 위해 추가

def show_margin_calc():
    st.title("🎯 페어리테이블 전략적 마진 시뮬레이션")
    st.markdown("---")

    # 1. 목표 수익률 설정 슬라이더
    target_margin_rate = st.sidebar.slider("🎯 목표 수익률 설정 (%)", 5, 50, 20) / 100
    st.sidebar.info(f"현재 목표 수익률 {int(target_margin_rate*100)}%를 기준으로 계산합니다.")

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        
        # 데이터 로드 (중략 - 이전 로직과 동일)
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        url_products = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}"
        url_costs = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}"
        
        df_products = pd.read_csv(url_products)
        df_costs = pd.read_csv(url_costs)
        
        # 전처리 및 월 선택
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()
        selected_month = st.selectbox("📅 분석할 원가 기준 월을 선택하세요", df_costs['월'].unique().tolist())

        # 원재료 단가 추출 (target_cost_row 정의)
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        sinjae = pd.to_numeric(str(target_cost_row.get('신재', 0)).replace(',', ''), errors='coerce')
        jaesaeng = pd.to_numeric(str(target_cost_row.get('재생', 0)).replace(',', ''), errors='coerce')
        anlyo_price = pd.to_numeric(str(target_cost_row.get('안료', 0)).replace(',', ''), errors='coerce')

        # 2. 핵심 계산 함수 (이전 공식 유지)
        def calc_logic(row):
            # ... (이전 코드의 calc_logic 내용과 동일) ...
            # [생략된 부분: 가로, 세로, 두께, 원단길이, 롤당수량, 박스비 활용 원가 계산 및 추천 납품가 역산]
            # (계산 결과 Series 반환)
            return pd.Series([total_cost, cur_nap_ga, rec_nap_ga, adjustment]) # 예시 반환값

        # 3. 결과 적용
        res_cols = ['제조 원가', '현재 납품가', '추천 납품가', '조정 필요액']
        # 실제 환경에서는 위 calc_logic 함수 내용이 전체 들어가야 합니다.
        df_res = df_products.apply(calc_logic, axis=1) # 여기서는 로직이 생략되었으므로 실제 코드 적용시 주의
        df_products[res_cols] = df_res
        display_df = df_products[['상품명'] + res_cols]

        # 4. 분석 결과 테이블 출력
        st.subheader(f"📊 {selected_month} 분석 결과")
        st.dataframe(display_df, use_container_width=True)

        # 5. [신규] 엑셀 다운로드 기능
        st.markdown("---")
        
        # 메모리상에 엑셀 파일 생성
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, index=False, sheet_name='마진분석결과')
        
        st.download_button(
            label="📥 분석 결과 엑셀로 다운로드",
            data=buffer.getvalue(),
            file_name=f"페어리테이블_마진분석_{selected_month}.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")

if check_password():
    show_margin_calc()
