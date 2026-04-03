import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V3.1)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 (Google Sheets)
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        sheet_name_3 = quote("월별납품가")
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        df_monthly_prices = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_3}")
        
        # 컬럼 공백 제거 및 SKU ID 포맷 통일
        for df in [df_products, df_costs, df_monthly_prices]:
            df.columns = df.columns.str.strip()
        
        def clean_sku(df):
            col = next((c for c in df.columns if 'SKU' in c.upper() or '상품코드' in c), df.columns[0])
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            return col

        prod_sku_col = clean_sku(df_products)
        monthly_sku_col = clean_sku(df_monthly_prices)

        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # 2. 사이드바 - 현재 월 자동 로드 및 일괄 조정
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 가격 조정 (%)", value=0.0, step=0.5)
        
        now = datetime.datetime.now()
        month_options = df_costs['월'].astype(str).str.strip().unique().tolist()
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
        
        # 원료 단가 추출 (유연한 검색)
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 월별 납품가 매핑
        price_col = selected_month if selected_month in df_monthly_prices.columns else '기본'
        master_price_dict = df_monthly_prices.set_index(monthly_sku_col)[price_col].to_dict()

        st.subheader("✍️ 납품 단가 시뮬레이션")
        st.info(f"✅ '{price_col}' 단가 리스트 적용 중")
        
        edit_df = df_products[[prod_sku_col, '상품명']].copy()
        edit_df.rename(columns={prod_sku_col: 'SKU ID'}, inplace=True)
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(master_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct / 100)).round(0).astype(int)
        
        edit_df.set_index('SKU ID', inplace=True)
        edited_output = st.data_editor(edit_df, use_container_width=True, key="v31_sync_editor")
        price_map = edited_output['적용납품가'].to_dict()

        # 4. 분석 로직 (V3.0 롤 무게 기반 계산식)
        def calc_logic(row):
            try:
                sku_id = row[prod_sku_col]
                applied_p = price_map.get(sku_id, 0)
                
                # 규격 파싱
                garo = clean_num(row.get('가로', 90))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # 원료비 계산
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r)

                # [표준 롤 무게 공식] 가로(cm) * 두께(mm) * 길이(m) * 0.0184
                roll_weight = garo * dukki * length * 0.0184
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                # 제조원가(박스) = (롤 전체 원료비 / 생산 박스 수) + 박스비
                total_box_cost = round(((roll_weight * unit_price) / boxes_per_roll) + box_cost, 0) if boxes_per_roll > 0 else 0
                
                # 추천가 및 수익
                rec_nap_ga = round(total_box_cost / 0.8, 0) # 목표 마진 20%
                current_roll_profit = (applied_p - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                return pd.Series([
                    sku_id, row.get('상품명', ''), f"{roll_weight:.2f}kg", total_box_cost, applied_p, 
                    rec_nap_ga, (applied_p - total_box_cost), current_roll_profit, status
                ], index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])
            except:
                return pd.Series([sku_id, '오류', '0kg', 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])

        # 5. 결과 출력
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 분석 리포트")
        
        df_disp = df_res.copy()
        for col in ['제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익']:
            df_disp[col] = df_disp[col].apply(lambda x: f"{int(x):,}원")

        st.dataframe(df_disp.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=['방어선']), use_container_width=True, hide_index=True)

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='분석리포트')
        st.download_button("📥 엑셀 다운로드", buffer.getvalue(), f"마진분석_{selected_month}.xlsx")

    except Exception as e:
        st.error(f"⚠️ 시스템 오류: {e}")

if check_password():
    show_margin_calc()
