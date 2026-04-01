import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.1 - 수정 기능 강화)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 및 전처리
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()

        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # 2. 원가 기준 선택
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # 원료 단가 추출 (제조원가 정상화 로직)
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # ---------------------------------------------------------
        # [핵심] 3. 데이터 유지(Session State) 로직
        # ---------------------------------------------------------
        # 시트의 납품가 컬럼 찾기
        orig_price_col = next((k for k in df_products.columns if '납품가' in k), None)

        # 수정한 단가를 세션에 저장 (초기화 방지)
        if 'working_df' not in st.session_state or st.session_state.get('current_month') != selected_month:
            temp_df = df_products[['SKU ID', '상품명']].copy()
            temp_df['적용납품가'] = df_products[orig_price_col].apply(clean_num)
            st.session_state.working_df = temp_df
            st.session_state.current_month = selected_month

        st.subheader("✍️ 납품 단가 시뮬레이션")
        st.info("💡 셀을 더블 클릭해 수정 후 반드시 'Enter'를 누르거나 표 바깥을 클릭하세요.")

        # 에디터에서 바로 세션 상태를 수정하도록 설정
        edited_df = st.data_editor(
            st.session_state.working_df,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(직접수정)", format="%d원")
            },
            hide_index=True, use_container_width=True, key="realtime_editor"
        )
        
        # 수정본을 다시 세션에 저장
        st.session_state.working_df = edited_df
        price_map = edited_df.set_index('SKU ID')['적용납품가'].to_dict()

        # ---------------------------------------------------------
        # 4. 분석 로직 (V2.1 계산식 고정 적용)
        # ---------------------------------------------------------
        COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS = \
            'SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '현재 상품수익', '현재 롤수익', '추천가 상품수익', '추천가 롤수익', '방어선'

        def calc_logic(row):
            try:
                sku_id = str(row.get('SKU ID', ''))
                # 세션에 저장된 수정 단가를 최우선 적용
                applied_p = price_map.get(sku_id, clean_num(row.get(orig_price_col, 0)))
                
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 1200))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                s_r, j_r, a_r = (v/100 if v > 1 else v for v in [s_val, j_val, a_val])
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                # 제조 원가 (비중 0.000184 고정)
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20))
                indiv_target = target_val / 100 if target_val > 1 else target_val
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                # 현재/추천 수익 계산
                current_unit_profit = applied_p - total_box_cost
                current_roll_profit = current_unit_profit * boxes_per_roll
                rec_unit_profit = rec_nap_ga - total_box_cost
                rec_roll_profit = rec_unit_profit * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id.split('.')[0], row.get('상품명', ''), fmt(total_box_cost), fmt(applied_p), 
                    fmt(rec_nap_ga), fmt(current_unit_profit), fmt(current_roll_profit),
                    fmt(rec_unit_profit), fmt(rec_roll_profit), status
                ], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS])
            except:
                return pd.Series([sku_id, '계산오류', '0원', '0원', '0원', '0원', '0원', '0원', '0원', '오류'], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS])

        # 5. 리포트 출력
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 실시간 분석 리포트")
        
        st.dataframe(
            df_res.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=[COL_STATUS]), 
            use_container_width=True, hide_index=True
        )

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
