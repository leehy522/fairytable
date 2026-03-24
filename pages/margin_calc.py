import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io

def show_margin_calc():
    st.title("🎯 페어리테이블 전략적 마진 시뮬레이션")
    st.markdown("---")

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
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                indiv_target = clean_num(row.get(target_col, 20))
                if indiv_target > 1: indiv_target /= 100 

                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                
                # [수정] 박스 1개당 들어가는 비닐 매수 (예: 100매)
                box_pcs = clean_num(row.get('매수', 100)) 
                box_cost = clean_num(row.get('박스비', 0))

                # [수정] 비율 데이터가 100인지 1.0인지 방어하는 로직
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                
                s_ratio = s_val / 100 if s_val > 1 else s_val
                j_ratio = j_val / 100 if j_val > 1 else j_val
                a_ratio = a_val / 100 if a_val > 1 else a_val

                # [핵심 수정] 1장 단위 정밀 계산법 도입
                # ① 비닐 1장의 무게 (kg) = 가로 * 세로 * 두께 * 0.000184 (2겹 비중 상수)
                single_weight = garo * sero * dukki * 0.000184
                
                # ② 1kg당 평균 재료 단가
                unit_price = (sinjae * s_ratio) + (jaesaeng * j_ratio) + (anlyo_price * a_ratio)
                
                # ③ 총 원가 = (1장 무게 * 1박스 매수 * kg당 단가) + 박스비
                total_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                rec_nap_ga = round(total_cost / (1 - indiv_target), 0)
                cur_nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                adjustment = rec_nap_ga - cur_nap_ga

                # [표기용 포맷팅] 천단위 콤마와 '원' 추가
                def fmt(v): return f"{int(v):,}원"
                
                return pd.Series([sku_val, row.get('상품명', ''), f"{int(indiv_target*100)}%", total_cost, cur_nap_ga, rec_nap_ga, adjustment])
            except Exception as e:
                # 에러 발생 시 로그를 남겨 디버깅 가능하도록 수정
                print(f"Error row {row.get('상품명')}: {e}")
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
                           file_name=f"페어리테이블_마진분석_{selected_month}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"오류 발생: {e}")

if check_password():
    show_margin_calc()
