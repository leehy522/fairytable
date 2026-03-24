import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io

def show_margin_calc():
    st.title("🎯 페어리테이블 상품별 전략 마진 분석")
    st.markdown("---")

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        url_products = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}"
        url_costs = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}"
        
        df_products = pd.read_csv(url_products)
        df_costs = pd.read_csv(url_costs)
        
        # 데이터 정리
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = str(value).replace(',', '').replace('%', '').replace('원', '').strip()
            return pd.to_numeric(s, errors='coerce')

        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo_price = clean_num(target_cost_row.get('안료', 0))

        def calc_logic(row):
            try:
                sku_val = str(row.get('SKU ID', '')).split('.')[0]
                
                # [개별 목표 수익률 가져오기]
                # 시트에 값이 없으면 기본값 20% (0.2) 적용
                indiv_target = clean_num(row.get('목표 수익률', 20))
                if indiv_target > 1: indiv_target /= 100 

                # 원가 계산 로직
                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                pcs_per_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                box_cost = clean_num(row.get('박스비', 0))

                s_ratio = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 0)) / 100
                j_ratio = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                a_ratio = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0)) / 100

                total_weight = garo * sero * dukki * 0.00000184 * length
                unit_price = (sinjae * s_ratio) + (jaesaeng * j_ratio) + (anlyo_price * a_ratio)
                total_cost = round(((total_weight * unit_price) / (pcs_per_roll if pcs_per_roll > 0 else 1)) + box_cost, 0)
                
                # 개별 수익률 적용 추천가 역산
                rec_nap_ga = round(total_cost / (1 - indiv_target), 0)
                cur_nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                adjustment = rec_nap_ga - cur_nap_ga
                
                return pd.Series([sku_val, row.get('상품명', ''), f"{int(indiv_target*100)}%", total_cost, cur_nap_ga, rec_nap_ga, adjustment])
            except:
                return pd.Series(['', row.get('상품명', ''), '20%', 0, 0, 0, 0])

        res_cols = ['SKU ID', '상품명', '설정 수익률', '제조 원가', '현재 납품가', '추천 납품가', '조정 필요액']
        df_res = df_products.apply(calc_logic, axis=1)
        df_products[res_cols] = df_res

        display_df = df_products[res_cols].copy()
        
        st.subheader(f"📊 {selected_month} 상품별 개별 마진 분석")
        
        def color_adj(val):
            if isinstance(val, (int, float)):
                return f'color: {"red" if val > 0 else "blue"}'
            return ''

        st.dataframe(display_df.style.applymap(color_adj, subset=['조정 필요액']), use_container_width=True)

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, index=False, sheet_name='상품별마진분석')
        
        st.download_button(label="📥 분석 결과 엑셀 다운로드", data=buffer.getvalue(), 
                           file_name=f"페어리테이블_개별마진_{selected_month}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"오류 발생: {e}")

if check_password():
    show_margin_calc()
