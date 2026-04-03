import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.6)")
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
        
        # [패치] 모든 칼럼명 전처리 (공백 및 특수기호 제거)
        for df in [df_products, df_costs, df_monthly_prices]:
            df.columns = df.columns.str.strip().str.replace(' ', '')

        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # [패치] SKU ID 칼럼 강제 지정 (첫 번째 열을 무조건 SKU로 간주)
        def get_sku_col(df):
            found = next((c for c in df.columns if 'SKU' in c.upper() or '상품코드' in c), df.columns[0])
            return found

        prod_sku = get_sku_col(df_products)
        monthly_sku = get_sku_col(df_monthly_prices)

        # SKU ID 문자열 통일 (.0 제거)
        df_products[prod_sku] = df_products[prod_sku].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_monthly_prices[monthly_sku] = df_monthly_prices[monthly_sku].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 사이드바 설정 (일괄 조정 및 월 선택)
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 가격 조정 (%)", value=0.0, step=0.5, help="시트 단가에 입력한 %만큼 가산합니다.")
        
        now = datetime.datetime.now()
        month_options = df_costs['월'].astype(str).str.strip().unique().tolist()
        
        # 현재 월(2026-04) 자동 선택 로직
        default_idx = 0
        for i, m in enumerate(month_options):
            if str(now.year) in m and str(now.month).zfill(2) in m:
                default_idx = i; break
            elif f"{now.month}월" in m:
                default_idx = i; break

        selected_month = st.selectbox("📅 원가 기준 월 선택", month_options, index=default_idx)
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # 원료 단가 추출
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))

        # ---------------------------------------------------------
        # 3. 월별/기본 납품가 매핑 (V2.3 로직)
        # ---------------------------------------------------------
        # 스크린샷의 '2026-04' 처럼 칼럼명이 존재하는지 확인
        current_price_col = selected_month.replace(' ', '')
        if current_price_col in df_monthly_prices.columns:
            price_col = current_price_col
            st.info(f"✅ '{price_col}' 전용 단가 리스트를 불러왔습니다.")
        else:
            price_col = '기본'
            st.info(f"ℹ️ '{selected_month}' 전용 단가가 없어 '기본' 단가를 적용합니다.")

        master_price_dict = df_monthly_prices.set_index(monthly_sku)[price_col].to_dict()

        st.subheader("✍️ 납품 단가 시뮬레이션")
        edit_df = df_products[[prod_sku, '상품명']].copy()
        edit_df.rename(columns={prod_sku: 'SKU ID'}, inplace=True)
        
        # 기초 단가 로드 + 일괄 조정(%) 반영
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(master_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct / 100)).round(0).astype(int)
        
        edit_df.set_index('SKU ID', inplace=True)

        # [실시간 동기화 락]
        edited_output = st.data_editor(
            edit_df,
            column_config={
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(수정)", format="%d원")
            },
            use_container_width=True, key="sync_editor_v26"
        )
        price_map = edited_output['적용납품가'].to_dict()

        # ---------------------------------------------------------
        # 4. 분석 로직 (V2.1 계산식 - 비중 0.000184 절대 보존)
        # ---------------------------------------------------------
        def calc_logic(row):
            try:
                sku_id = row[prod_sku]
                applied_p = price_map.get(sku_id, 0)
                
                # 규격 데이터
                garo = clean_num(row.get('가로', 90)); sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100)); box_cost = clean_num(row.get('박스비', 0))
                
                # 원재료비 산출
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r)

                # [공식] 제조원가 = ((가로 * 세로 * 두께 * 0.000184) * 매수 * 단위단가) + 박스비
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                # 추천가 (목표마진 20% 기준)
                rec_nap_ga = round(total_box_cost / 0.8, 0)
                
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                # 수익 지표
                current_roll_profit = (applied_p - total_box_cost) * boxes_per_roll
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                return pd.Series([
                    sku_id, row.get('상품명', ''), total_box_cost, applied_p, 
                    rec_nap_ga, (applied_p - total_box_cost), current_roll_profit, status
                ], index=['SKU ID', '상품명', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])
            except:
                return pd.Series([sku_id, '계산오류', 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])

        # 5. 리포트 출력 및 엑셀 다운로드
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 분석 리포트")
        
        # 화면 출력용 포맷팅
        df_disp = df_res.copy()
        for c in ['제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익']:
            df_disp[c] = df_disp[c].apply(lambda x: f"{int(x):,}원")

        st.dataframe(
            df_disp.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=['방어선']), 
            use_container_width=True, hide_index=True
        )

        # [신규] 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='분석결과')
        st.download_button("📥 엑셀 리포트 다운로드", buffer.getvalue(), f"마진분석_{selected_month}.xlsx")

    except Exception as e:
        st.error(f"⚠️ 시스템 오류: {e}")
        st.info("데이터프레임 칼럼 확인: " + str(df_monthly_prices.columns.tolist() if 'df_monthly_prices' in locals() else "로드실패"))

if check_password():
    show_margin_calc()
