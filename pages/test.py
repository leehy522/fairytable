import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import os

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.2 - 추천가 방어선 추가)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 (Google Sheets)
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # 2. 원가 기준 설정
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 납품가 편집기
        st.subheader("✍️ 납품 단가 시뮬레이션")
        orig_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_df = df_products[['SKU ID', '상품명']].copy()
        edit_df['적용납품가'] = df_products[orig_price_col].apply(clean_num)

        edited_output = st.data_editor(
            edit_df,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(직접수정)", format="%d원")
            },
            hide_index=True, use_container_width=True, key="realtime_editor"
        )
        price_map = edited_output.set_index('SKU ID')['적용납품가'].to_dict()

        # 4. 분석 로직 (기존 계산식 보존 + 추천가 방어선 추가)
        COL_SKU = 'SKU ID'
        COL_NAME = '상품명'
        COL_COST = '제조원가(박스)'
        COL_APPLIED = '적용납품가'
        COL_REC = '추천납품가'
        COL_PROFIT = '현재 롤수익'
        COL_STATUS = '현재 방어선'
        COL_REC_PROFIT = '추천가 롤수익'
        COL_REC_STATUS = '추천가 방어선' # 새로 추가된 항목

        def get_status(profit):
            """수익에 따른 방어선 상태 결정"""
            if profit < 15000: return "🚨 적자위험"
            elif profit > 50000: return "⚠️ 고마진"
            return "✅ 정상"

        def calc_logic(row):
            try:
                sku_id = str(row.get('SKU ID', ''))
                applied_p = price_map.get(sku_id, clean_num(row.get(orig_price_col, 0)))
                
                # 규격 데이터
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # 단위 원가 계산
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r)

                # 원가 및 추천가 계산
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20)) / 100
                rec_nap_ga = round(total_box_cost / (1 - target_val), 0)
                
                # 롤당 박스 수량
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                # 현재 수익 및 방어선
                current_roll_profit = (applied_p - total_box_cost) * boxes_per_roll
                current_status = get_status(current_roll_profit)
                
                # 추천가 수익 및 방어선
                rec_roll_profit = (rec_nap_ga - total_box_cost) * boxes_per_roll
                rec_status = get_status(rec_roll_profit)

                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id.split('.')[0], row.get('상품명', ''), fmt(total_box_cost), fmt(applied_p), 
                    fmt(rec_nap_ga), fmt(current_roll_profit), current_status,
                    fmt(rec_roll_profit), rec_status
                ], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_PROFIT, COL_STATUS, COL_REC_PROFIT, COL_REC_STATUS])
            except:
                return pd.Series([sku_id, '오류', '0원', '0원', '0원', '0원', '오류', '0원', '오류'], 
                                 index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_PROFIT, COL_STATUS, COL_REC_PROFIT, COL_REC_STATUS])

        # 5. 결과 리포트 출력
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 및 추천가 방어선 분석")
        
        def style_status(val):
            if '🚨' in str(val): return 'color: #d32f2f; font-weight: bold;'
            if '⚠️' in str(val): return 'color: #f57c00;'
            return ''

        st.dataframe(
            df_res.style.map(style_status, subset=[COL_STATUS, COL_REC_STATUS]), 
            use_container_width=True, 
            hide_index=True
        )

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
