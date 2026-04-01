import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.0 - 계산식 고정)")
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
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            # 숫자와 마침표만 남기고 모두 제거
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # 2. 원가 기준 설정
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # [수정] 원단값이 0원으로 나오지 않도록 추출 로직을 안정화했습니다.
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        # 재생 단가 추출 시 발생하던 제너레이터 오류를 해결했습니다.
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

        # 4. 분석 로직 (V2.0 사용자 계산식 절대 고정)
        COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS = \
            'SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '단가 조정액(+/-)', '쿠팡판매가(42%)', '롤당수익', '방어선'

        def calc_logic(row):
            try:
                sku_id = str(row.get('SKU ID', ''))
                applied_p = price_map.get(sku_id, clean_num(row.get(orig_price_col, 0)))
                
                # [유지] 규격 데이터
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # [유지] 배합비 및 단위 원가
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r)

                # [유지] 무게 및 원가 계산
                effective_sero = sero if sero > 0 else (length / box_pcs if box_pcs > 0 else 1)
                single_weight = garo * effective_sero * dukki * 0.000184
                vinyl_cost = single_weight * box_pcs * unit_price
                total_box_cost = round(vinyl_cost + box_cost, 0)
                
                # [유지] 추천납품가 (목표수익률 반영)
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_margin = clean_num(row.get(target_col, 20)) / 100
                rec_nap_ga = round(total_box_cost / (1 - target_margin), 0)
                
                # [유지] 지표 산출
                adj_val = rec_nap_ga - applied_p
                coupang_selling = round(applied_p / 0.58, 0)
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                current_roll_profit = (applied_p - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"
                def fmt_adj(v): return f"{'+' if v > 0 else ''}{int(round(v, 0)):,}원"

                return pd.Series([sku_id.split('.')[0], row.get('상품명', ''), fmt(total_box_cost), fmt(applied_p), fmt(rec_nap_ga), fmt_adj(adj_val), fmt(coupang_selling), fmt(current_roll_profit), status],
                                 index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])
            except:
                return pd.Series([sku_id, '계산오류', '0원', '0원', '0원', '0원', '0원', '0원', '오류'], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])

        # 5. 리포트 출력 및 스타일링
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 분석 리포트")
        
        # [수정] 최신 Pandas 버전 에러 대응 (applymap 대신 map 사용)
        st.dataframe(
            df_res.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=[COL_STATUS]), 
            use_container_width=True, 
            hide_index=True
        )

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
