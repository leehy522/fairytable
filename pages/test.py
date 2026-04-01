import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import os

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 실시간 분석기 (V3.0)")
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
        
        # [중요] 원료 단가 추출 로직 보강 (제조원가 정상화의 핵심)
        sinjae_price = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng_price = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo_price = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 납품가 수동 편집 섹션 (실시간 반영 엔진)
        st.subheader("✍️ 납품 단가 시뮬레이션")
        st.caption("아래 표의 '적용납품가'를 수정하면 하단 리포트가 즉시 갱신됩니다.")
        
        orig_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_base = df_products[['SKU ID', '상품명']].copy()
        edit_base['적용납품가'] = df_products[orig_price_col].apply(clean_num)

        # 데이터 에디터 (key를 사용하여 변화 감지)
        edited_price_df = st.data_editor(
            edit_base,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(수정)", format="%d원")
            },
            hide_index=True,
            use_container_width=True,
            key="realtime_margin_editor"
        )
        
        # 수정된 단가를 매핑
        price_map = edited_price_df.set_index('SKU ID')['적용납품가'].to_dict()

        # 4. 분석 로직 (실시간 데이터 반영)
        COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_COUPANG, COL_STATUS = \
            'SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '쿠팡가(42%)', '방어선'

        def calc_logic(row):
            try:
                sku_id = str(row.get('SKU ID', ''))
                # 수정한 단가 적용, 없으면 기존 단가
                applied_p = price_map.get(sku_id, clean_num(row.get(orig_price_col, 0)))
                
                # 규격 및 원가 계산
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae_price * s_r) + (jaesaeng_price * j_r)

                # 원가 계산 (비닐값 + 박스비)
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                # 추천가 및 수익 지표
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20)) / 100
                rec_nap_ga = round(total_box_cost / (1 - target_val), 0)
                
                unit_profit = applied_p - total_box_cost
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                roll_profit = unit_profit * boxes_per_roll
                
                coupang_price = round(applied_p / 0.58, 0)
                
                status = "✅ 정상"
                if roll_profit < 15000: status = "🚨 적자위험"
                elif roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id.split('.')[0], row.get('상품명', ''), fmt(total_box_cost), fmt(applied_p), 
                    fmt(rec_nap_ga), fmt(unit_profit), fmt(roll_profit), fmt(coupang_price), status
                ], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_COUPANG, COL_STATUS])
            except:
                return pd.Series([sku_id, '계산오류', '0원', '0원', '0원', '0원', '0원', '0원', '오류'], 
                                 index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_COUPANG, COL_STATUS])

        # 5. 결과 리포트 출력
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 실시간 분석 리포트")
        
        # 스타일 적용 (🚨 적자위험 시 빨간색 강조)
        def style_status(val):
            if '🚨' in str(val): return 'color: #d32f2f; font-weight: bold;'
            if '⚠️' in str(val): return 'color: #f57c00;'
            return ''

        st.dataframe(
            df_res.style.map(style_status, subset=[COL_STATUS]), 
            use_container_width=True, 
            hide_index=True
        )

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
