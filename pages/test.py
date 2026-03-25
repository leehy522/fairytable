import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re

def show_margin_calc():
    st.title("🛡️ 페어리테이블 통합 마진 분석 및 가격 방어 시스템")
    st.markdown("---")

    try:
        # 1. 데이터 로드
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo = clean_num(target_cost_row.get('안료', 0))

        def calc_logic(row):
            try:
                # [기본 데이터]
                sku_val = str(row.get('SKU ID', '')).split('.')[0]
                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 1200))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                cur_nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                
                # 목표 수익률 설정
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                indiv_target = clean_num(row.get(target_col, 20))
                if indiv_target > 1: indiv_target /= 100

                # 배합비 계산
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                s_r, j_r, a_r = (v/100 if v > 1 else v for v in [s_val, j_val, a_val])
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                # --- [A. 원단 롤(Roll) 분석] ---
                # 롤 무게 (폭*두께*길이*0.0184)
                roll_weight = garo * dukki * length * 0.0184
                roll_material_cost = roll_weight * unit_price
                
                # --- [B. 박스 단위 분석] ---
                # 비닐 1장 무게 및 박스 제조원가
                single_weight = garo * sero * dukki * 0.000184
                box_material_cost = single_weight * box_pcs * unit_price
                total_box_cost = round(box_material_cost + box_cost, 0)
                
                # 추천 납품가 산출
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                adjustment = rec_nap_ga - cur_nap_ga

                # --- [C. 가격 방어선 알고리즘] ---
                # 롤당 최소 가공 수익 방어 (예: 25,000원)
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                current_roll_profit = (cur_nap_ga - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                # 포맷팅 함수
                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_val, row.get('상품명', ''), f"{roll_weight:.2f}kg",
                    fmt(total_box_cost), fmt(cur_nap_ga), fmt(rec_nap_ga), fmt(adjustment),
                    fmt(current_roll_profit), status
                ])
            except:
                return pd.Series(['', row.get('상품명', ''), "0kg", "0원", "0원", "0원", "0원", "0원", "오류"])

        res_cols = ['SKU ID', '상품명', '롤무게', '제조원가(박스)', '현재납품가', '추천납품가', '조정필요액', '롤당수익', '방어선']
        df_res = df_products.apply(calc_logic, axis=1)
        df_products[res_cols] = df_res

        # UI 출력 및 스타일링
        st.subheader(f"📊 {selected_month} 통합 분석 리포트 (부가세 별도)")
        
        def style_logic(val):
            if '🚨' in str(val): return 'color: red; font-weight: bold'
            if '⚠️' in str(val): return 'color: orange'
            return ''

        st.dataframe(df_products[res_cols].style.applymap(style_logic, subset=['방어선']), use_container_width=True)

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_products[res_cols].to_excel(writer, index=False, sheet_name='통합분석')
        st.download_button("📥 분석 결과 엑셀 다운로드", buffer.getvalue(), f"페어리테이블_통합분석_{selected_month}.xlsx")

    except Exception as e:
        st.error(f"실행 오류: {e}")

if check_password():
    show_margin_calc()
