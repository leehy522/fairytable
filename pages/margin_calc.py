import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io

def show_margin_calc():
    st.title("🎯 페어리테이블 전략적 마진 시뮬레이션")
    st.markdown("---")

    # 1. 목표 수익률 설정 슬라이더
    target_margin_rate = st.sidebar.slider("🎯 목표 수익률 설정 (%)", 5, 50, 20) / 100
    st.sidebar.info(f"현재 목표 수익률 {int(target_margin_rate*100)}%를 기준으로 계산합니다.")

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        url_products = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}"
        url_costs = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}"
        
        df_products = pd.read_csv(url_products)
        df_costs = pd.read_csv(url_costs)
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        selected_month = st.selectbox("📅 분석할 원가 기준 월을 선택하세요", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = str(value).replace(',', '').replace('%', '').replace('원', '').strip()
            return pd.to_numeric(s, errors='coerce')

        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo_price = clean_num(target_cost_row.get('안료', 0))

        # 2. 핵심 계산 함수
        def calc_logic(row):
            try:
                # [SKU ID + 상품명 결합]
                sku_id = str(row.get('SKU ID', '')).split('.')[0] # 소수점 제거
                p_name = f"[{sku_id}] {row.get('상품명', '')}"
                
                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                pcs_per_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                box_cost = clean_num(row.get('박스비', 0))

                s_ratio = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 0))
                j_ratio = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_ratio = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                
                if s_ratio > 1: s_ratio /= 100
                if j_ratio > 1: j_ratio /= 100
                if a_ratio > 1: a_ratio /= 100

                # 공식 적용
                total_weight = garo * sero * dukki * 0.00000184 * length
                unit_price = (sinjae * s_ratio) + (jaesaeng * j_ratio) + (anlyo_price * a_ratio)
                total_cost = round(((total_weight * unit_price) / (pcs_per_roll if pcs_per_roll > 0 else 1)) + box_cost, 0)
                
                rec_nap_ga = round(total_cost / (1 - target_margin_rate), 0)
                cur_nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                adjustment = rec_nap_ga - cur_nap_ga
                
                return pd.Series([p_name, total_cost, cur_nap_ga, rec_nap_ga, adjustment])
            except:
                return pd.Series([row.get('상품명', ''), 0, 0, 0, 0])

        # 3. 결과 적용
        res_cols = ['표시상품명', '제조 원가', '현재 납품가', '추천 납품가', '조정 필요액']
        df_res = df_products.apply(calc_logic, axis=1)
        df_products[res_cols] = df_res

        # 분석 결과 테이블
        display_df = df_products[res_cols].copy()
        display_df.columns = ['상품명(SKU 포함)', '제조 원가', '현재 납품가', '추천 납품가', '조정 필요액']
        
        st.subheader(f"📊 {selected_month} 분석 결과 (목표 {int(target_margin_rate*100)}%)")
        
        def color_adj(val):
            color = 'red' if val > 0 else 'blue'
            return f'color: {color}'

        st.dataframe(display_df.style.applymap(color_adj, subset=['조정 필요액']), use_container_width=True)

        # 4. 엑셀 다운로드
        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, index=False, sheet_name='마진분석')
        
        st.download_button(
            label="📥 분석 결과 엑셀로 다운로드",
            data=buffer.getvalue(),
            file_name=f"페어리테이블_SKU별_분석_{selected_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")

if check_password():
    show_margin_calc()
