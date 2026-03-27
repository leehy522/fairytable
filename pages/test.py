import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 실시간 시뮬레이터")
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
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # 2. 원가 기준 선택
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo = clean_num(target_cost_row.get('안료', 0))

        # ---------------------------------------------------------
        # [핵심 수정] 3. 세션 상태를 이용한 실시간 단가 매핑
        # ---------------------------------------------------------
        st.subheader("✍️ 납품 단가 수정 (입력 즉시 하단 반영)")
        
        # 기본 단가 데이터프레임 준비
        original_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_df = df_products[['SKU ID', '상품명']].copy()
        edit_df['수정납품가'] = df_products[original_price_col].apply(clean_num)

        # 데이터 에디터 (key를 지정하여 변화를 감지)
        edited_output = st.data_editor(
            edit_df,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "수정납품가": st.column_config.NumberColumn("납품가(수정)", format="%d원")
            },
            hide_index=True,
            use_container_width=True,
            key="price_editor"
        )
        
        # 수정된 값을 딕셔너리로 변환 (하단 계산기에서 참조)
        current_prices = edited_output.set_index('SKU ID')['수정납품가'].to_dict()

        # ---------------------------------------------------------
        # 4. 분석 로직 (수정된 단가 current_prices를 직접 참조)
        # ---------------------------------------------------------
        COL_SKU = 'SKU ID'
        COL_NAME = '상품명'
        COL_BOX_COST = '제조원가'
        COL_APPLIED_PRICE = '적용납품가' # 여기가 즉시 바뀌어야 함
        COL_COUPANG = '쿠팡판매가(42%)'
        COL_PROFIT = '롤당수익'
        COL_STATUS = '방어선'

        def calc_logic(row):
            try:
                sku_id = str(row.get('SKU ID', ''))
                # 편집기에서 수정한 값을 가져오고, 없으면 시트 값 사용
                applied_price = current_prices.get(sku_id, clean_num(row.get(original_price_col, 0)))
                
                # 규격 및 무게 계산
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                a_r = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                # [즉시 반영될 지표들]
                coupang_price = round(applied_price / 0.58, 0)
                
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                current_roll_profit = (applied_price - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id.split('.')[0],
                    row.get('상품명', ''),
                    fmt(total_box_cost),
                    fmt(applied_price), # 실시간 반영 확인용
                    fmt(coupang_price),
                    fmt(current_roll_profit),
                    status
                ], index=[COL_SKU, COL_NAME, COL_BOX_COST, COL_APPLIED_PRICE, COL_COUPANG, COL_PROFIT, COL_STATUS])
            except:
                return pd.Series(['', '', '0원', '0원', '0원', '0원', '오류'], index=[COL_SKU, COL_NAME, COL_BOX_COST, COL_APPLIED_PRICE, COL_COUPANG, COL_PROFIT, COL_STATUS])

        # 리포트 출력
        df_res = df_products.apply(calc_logic, axis=1)
        
        st.subheader("📊 시뮬레이션 결과 리포트")
        
        def style_row(val):
            if '🚨' in str(val): return 'background-color: #ffebee; color: #c62828;'
            return ''

        st.dataframe(
            df_res.style.applymap(style_row, subset=[COL_STATUS]),
            use_container_width=True,
            hide_index=True
        )

    except Exception as e:
        st.error(f"오류: {e}")

if check_password():
    show_margin_calc()
