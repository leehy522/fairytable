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
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        selected_month = st.selectbox("📅 분석 원가 기준 월", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            return pd.to_numeric(str(value).replace(',', '').replace('%', '').replace('원', '').strip(), errors='coerce')

        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo_p = clean_num(target_cost_row.get('안료', 0))

        def calc_logic(row):
            try:
                # 1. 규격 및 단가 정보
                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                
                # [수정포인트] 롤당 총 매수와 박스당 매수를 명확히 구분
                total_pcs = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                box_pcs = clean_num(row.get('매수', 100)) # 한 박스에 들어가는 장수 (예: 100매)
                box_cost = clean_num(row.get('박스비', 0))

                s_ratio = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_ratio = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                a_ratio = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0)) / 100

                # 2. [논리 수정] 낱장 단가 기반 계산
                # ① 롤 전체 무게 (kg)
                total_weight = garo * sero * dukki * 0.00000184 * length
                # ② 1kg당 평균 재료 단가
                avg_unit_price = (sinjae * s_ratio) + (jaesaeng * j_ratio) + (anlyo_p * a_ratio)
                # ③ 롤 전체 재료비
                total_material_cost = total_weight * avg_unit_price
                
                # ④ [핵심] 낱장당 가격 구하기 (롤 전체 가격 / 롤당 총 매수)
                single_pc_price = total_material_cost / (total_pcs if total_pcs > 0 else 1)
                
                # ⑤ [최종] (낱장 가격 * 박스당 매수) + 박스비
                total_cost = round((single_pc_price * box_pcs) + box_cost, 0)

                # 3. 목표 수익률 및 추천가 계산
                target_rate = clean_num(next((row[k] for k in row.index if '목표' in k and '율' in k), 20))
                if target_rate > 1: target_rate /= 100
                
                rec_nap_ga = round(total_cost / (1 - target_rate), 0)
                cur_nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                
                return pd.Series([row.get('SKU ID', ''), row.get('상품명', ''), f"{int(target_rate*100)}%", total_cost, cur_nap_ga, rec_nap_ga])
            except:
                return pd.Series(['', row.get('상품명', ''), '20%', 0, 0, 0])

        res_cols = ['SKU ID', '상품명', '목표수익률', '제조 원가', '현재 납품가', '추천 납품가']
        df_res = df_products.apply(calc_logic, axis=1)
        df_products[res_cols] = df_res
        
        st.dataframe(df_products[res_cols], use_container_width=True)

        # 엑셀 다운로드 (생략 - 이전과 동일)
    except Exception as e:
        st.error(f"오류: {e}")

if check_password():
    show_margin_calc()
