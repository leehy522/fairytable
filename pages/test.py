import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re

def show_margin_calc():
    st.title("🛡️ 페어리테이블 통합 마진 분석 시스템")
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

        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo = clean_num(target_cost_row.get('안료', 0))

        # [표기용 이름 정의] - 여기서 정의한 이름과 return하는 데이터 순서가 완벽히 일치해야 합니다.
        COL_SKU = 'SKU ID'
        COL_NAME = '상품명'
        COL_ROLL_W = '롤무게'
        COL_BOX_COST = '제조원가(원)'
        COL_CUR_PRICE = '현재납품가(원)'
        COL_REC_PRICE = '추천납품가(원)'
        COL_ADJ = '단가 조정액(+/-)'
        COL_COUPANG = '쿠팡판매가(42%)' # 새로 추가된 열
        COL_PROFIT = '롤당수익'
        COL_STATUS = '방어선'

        def calc_logic(row):
            try:
                # [기본 계산 데이터]
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 1200))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                cur_nap_ga = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                
                # [목표 수익률]
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20))
                indiv_target = target_val / 100 if target_val > 1 else target_val

                # [단가 계산]
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                s_r, j_r, a_r = (v/100 if v > 1 else v for v in [s_val, j_val, a_val])
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                # [박스 원가 및 추천가]
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                adjustment_val = rec_nap_ga - cur_nap_ga

                # [기존] 판매가 대비 수익률 42% 보전 (Margin)
                # 판매가 = 납품가 / 0.58
                coupang_selling_price = round(cur_nap_ga / 0.58, 0)

                # [롤당 수익 및 방어선]
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                current_roll_profit = (cur_nap_ga - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"
                def fmt_adj(v): return f"{'+' if v > 0 else ''}{int(round(v, 0)):,}원"

                return pd.Series([
                    str(row.get('SKU ID', '')).split('.')[0],
                    row.get('상품명', ''),
                    fmt(total_box_cost),
                    fmt(cur_nap_ga),
                    fmt(rec_nap_ga),
                    fmt_adj(adjustment_val),
                    fmt(coupang_selling_price), # 쿠팡 판매가 추가
                    fmt(current_roll_profit),
                    status
                ], index=[COL_SKU, COL_NAME, COL_BOX_COST, COL_CUR_PRICE, COL_REC_PRICE, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])
            except:
                return pd.Series(['', row.get('상품명', ''), '0원', '0원', '0원', '0원', '0원', '0원', '오류'], 
                                 index=[COL_SKU, COL_NAME, COL_BOX_COST, COL_CUR_PRICE, COL_REC_PRICE, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])

        df_res = df_products.apply(calc_logic, axis=1)
        display_cols = [COL_SKU, COL_NAME, COL_BOX_COST, COL_CUR_PRICE, COL_REC_PRICE, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS]
        
        st.subheader(f"📊 {selected_month} 쿠팡 납품 및 판매가 분석 리포트")
        
        def style_status(val):
            if '🚨' in str(val): return 'color: red; font-weight: bold'
            if '⚠️' in str(val): return 'color: orange'
            return ''

        st.dataframe(df_res[display_cols].style.applymap(style_status, subset=[COL_STATUS]), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res[display_cols].to_excel(writer, index=False, sheet_name='통합분석')
        st.download_button("📥 엑셀 다운로드", buffer.getvalue(), f"페어리테이블_쿠팡분석_{selected_month}.xlsx")

    except Exception as e:
        st.error(f"실행 오류: {e}")

if check_password():
    show_margin_calc()
