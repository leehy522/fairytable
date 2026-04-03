import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.4)")
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
        
        # [패치] 모든 데이터프레임의 컬럼명 공백 제거 및 인식 강화
        for df in [df_products, df_costs, df_monthly_prices]:
            df.columns = df.columns.str.strip()

        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # [패치] SKU ID 컬럼을 더 유연하게 찾습니다 (오타 대비)
        def find_sku_col(df):
            return next((c for c in df.columns if 'SKU' in c.upper() or '상품코드' in c), None)

        prod_sku_col = find_sku_col(df_products)
        monthly_sku_col = find_sku_col(df_monthly_prices)

        if not monthly_sku_col:
            st.error("❌ '월별납품가' 시트에서 'SKU ID' 컬럼을 찾을 수 없습니다. 시트의 첫 줄 헤더를 확인해주세요.")
            return

        # SKU ID 포맷 통일
        df_products[prod_sku_col] = df_products[prod_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_monthly_prices[monthly_sku_col] = df_monthly_prices[monthly_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 사이드바 - 월 선택 및 일괄 조정
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
        
        # 원료 단가 추출
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 월별 납품가 매핑
        if selected_month in df_monthly_prices.columns:
            price_col = selected_month
            st.info(f"✅ 현재 '{selected_month}' 특수 단가를 적용 중입니다.")
        else:
            price_col = '기본'
            st.info(f"ℹ️ '{selected_month}' 전용 단가가 없어 '기본' 단가를 적용합니다.")

        # [패치] 유연하게 찾은 SKU ID 컬럼을 기준으로 맵 생성
        master_price_dict = df_monthly_prices.set_index(monthly_sku_col)[price_col].to_dict()

        st.subheader("✍️ 납품 단가 시뮬레이션")
        edit_df = df_products[[prod_sku_col, '상품명']].copy()
        edit_df.rename(columns={prod_sku_col: 'SKU ID'}, inplace=True) # 표시용 이름 통일
        
        def get_base_price(sku):
            return clean_num(master_price_dict.get(sku, 0))

        edit_df['적용납품가'] = edit_df['SKU ID'].apply(get_base_price)
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct / 100)).round(0).astype(int)
        
        edit_df.set_index('SKU ID', inplace=True)

        edited_output = st.data_editor(
            edit_df,
            column_config={
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(직접수정 가능)", format="%d원")
            },
            use_container_width=True, 
            key="realtime_sync_editor_v10"
        )
        
        price_map = edited_output['적용납품가'].to_dict()

        # 4. 분석 로직 (V2.1 계산식 절대 고정)
        def calc_logic(row):
            try:
                sku_id = row[prod_sku_col]
                applied_p = price_map.get(sku_id, 0)
                
                # 규격 데이터 (비중 0.000184 고정)
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 1200))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r)

                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20))
                indiv_target = target_val / 100 if target_val > 1 else target_val
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                current_roll_profit = (applied_p - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                return pd.Series([
                    sku_id, row.get('상품명', ''), total_box_cost, applied_p, 
                    rec_nap_ga, (applied_p - total_box_cost), current_roll_profit,
                    (rec_nap_ga - total_box_cost), (rec_nap_ga - total_box_cost) * boxes_per_roll, status
                ], index=['SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '현재 상품수익', '현재 롤수익', '추천가 상품수익', '추천가 롤수익', '방어선'])
            except Exception as e:
                return pd.Series([sku_id, '계산오류', 0, 0, 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '현재 상품수익', '현재 롤수익', '추천가 상품수익', '추천가 롤수익', '방어선'])

        # 5. 출력
        df_res = df_products.apply(calc_logic, axis=1)
        
        df_display = df_res.copy()
        for col in ['제조원가(박스)', '적용납품가', '추천납품가', '현재 상품수익', '현재 롤수익', '추천가 상품수익', '추천가 롤수익']:
            df_display[col] = df_display[col].apply(lambda x: f"{int(round(x, 0)):,}원")

        st.dataframe(
            df_display.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=['방어선']), 
            use_container_width=True, hide_index=True
        )

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='분석리포트')
        st.download_button("📥 엑셀 다운로드", buffer.getvalue(), f"마진분석_{selected_month}.xlsx")

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
