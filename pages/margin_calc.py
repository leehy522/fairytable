import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.2)")
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

        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        df_products['SKU ID'] = df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # ---------------------------------------------------------
        # 2. 사이드바 - 일괄 가격 조정 (%) 및 월 선택
        # ---------------------------------------------------------
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        
        # [신규] 퍼센트 단위 일괄 조정 기능
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 가격 조정 (%)", value=0.0, step=0.5, help="시트의 원래 납품가에서 입력한 %만큼 가산하여 시작합니다.")
        
        now = datetime.datetime.now()
        month_options = df_costs['월'].unique().tolist()
        default_month_index = 0
        for i, m_val in enumerate(month_options):
            if str(now.year) in m_val and str(now.month).zfill(2) in m_val:
                default_month_index = i
                break
            elif f"{now.month}월" in m_val:
                default_month_index = i
                break

        selected_month = st.selectbox("📅 원가 기준 월 선택", month_options, index=default_month_index)
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 납품가 편집기 (일괄 조정 반영)
        st.subheader("✍️ 납품 단가 시뮬레이션")
        if adj_pct != 0:
            st.warning(f"💡 현재 모든 품목에 {adj_pct}% 인상/인하가 적용된 상태입니다.")
            
        orig_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_df = df_products[['SKU ID', '상품명']].copy()
        
        # [핵심] 일괄 조정 로직: (원본가) * (1 + 조정률/100)
        base_prices = df_products[orig_price_col].apply(clean_num)
        edit_df['적용납품가'] = (base_prices * (1 + adj_pct / 100)).round(0).astype(int)
        
        edit_df.set_index('SKU ID', inplace=True)

        edited_output = st.data_editor(
            edit_df,
            column_config={
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(직접수정 가능)", format="%d원")
            },
            use_container_width=True, 
            key="realtime_sync_editor_v8"
        )
        
        price_map = edited_output['적용납품가'].to_dict()

        # 4. 분석 로직 (V2.1 계산식 절대 보존)
        COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS = \
            'SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '현재 상품수익', '현재 롤수익', '추천가 상품수익', '추천가 롤수익', '방어선'

        def calc_logic(row):
            try:
                sku_id = row['SKU ID']
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

                # [V2.1 표준 수식] 0.000184 비중 고정
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20))
                indiv_target = target_val / 100 if target_val > 1 else target_val
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                current_unit_profit = applied_p - total_box_cost
                current_roll_profit = current_unit_profit * boxes_per_roll
                rec_unit_profit = rec_nap_ga - total_box_cost
                rec_roll_profit = rec_unit_profit * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                return pd.Series([
                    sku_id, row.get('상품명', ''), total_box_cost, applied_p, 
                    rec_nap_ga, current_unit_profit, current_roll_profit,
                    rec_unit_profit, rec_roll_profit, status
                ], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS])
            except:
                return pd.Series([row.get('SKU ID', ''), '오류', 0, 0, 0, 0, 0, 0, 0, '오류'], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS])

        # 5. 리포트 출력 및 엑셀 다운로드
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 상세 분석")
        
        # 화면 출력용 포맷팅 (쉼표 및 원 표시)
        df_display = df_res.copy()
        for col in [COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT]:
            df_display[col] = df_display[col].apply(lambda x: f"{int(x):,}원")

        st.dataframe(
            df_display.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=[COL_STATUS]), 
            use_container_width=True, hide_index=True
        )

        # [신규] 엑셀 다운로드 버튼
        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='마진분석리포트')
            # 엑셀 서식 자동 조정 등 추가 가능
        
        st.download_button(
            label="📥 분석 리포트 엑셀 다운로드",
            data=buffer.getvalue(),
            file_name=f"페어리테이블_마진분석_{selected_month}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
