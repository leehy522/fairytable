import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import os
import datetime # [추가] 오늘 날짜 인식을 위한 모듈

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V2.1 - 자동 월 로드)")
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

        # [기준 준수] SKU ID 완벽 통일 (문자열 클렌징)
        df_products['SKU ID'] = df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # ---------------------------------------------------------
        # [핵심] 2. 현재 월 자동 기본값 설정 (market_calc 동일 적용)
        # ---------------------------------------------------------
        now = datetime.datetime.now()
        month_options = df_costs['월'].unique().tolist()
        
        # 기본 인덱스 설정 (찾지 못할 경우 0번째 항목)
        default_month_index = 0
        
        for i, m_val in enumerate(month_options):
            # '2026-04' 또는 '4월' 등의 형식을 오늘 날짜와 대조
            if str(now.year) in m_val and str(now.month).zfill(2) in m_val:
                default_month_index = i
                break
            elif f"{now.month}월" in m_val:
                default_month_index = i
                break

        selected_month = st.selectbox(
            "📅 원가 기준 월 선택", 
            month_options, 
            index=default_month_index # 현재 월이 자동으로 먼저 선택됨
        )
        # ---------------------------------------------------------

        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # 원료 단가 추출 (안정화 로직)
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 납품가 편집기 (V2.1 완벽 동기화 표준)
        st.subheader("✍️ 납품 단가 시뮬레이션")
        st.info("💡 단가 숫자 입력 후 **반드시 키보드의 'Enter(엔터)' 키를 누르거나 표 밖을 클릭**해야 적용됩니다.")
        
        orig_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_df = df_products[['SKU ID', '상품명']].copy()
        edit_df['적용납품가'] = df_products[orig_price_col].apply(clean_num)
        
        # [기준 준수] SKU ID를 인덱스로 고정하여 데이터 유실 차단
        edit_df.set_index('SKU ID', inplace=True)

        edited_output = st.data_editor(
            edit_df,
            column_config={
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(직접수정)", format="%d원")
            },
            use_container_width=True, 
            key="realtime_sync_editor_v7" # 캐시 충돌 방지를 위해 키 버전업
        )
        
        price_map = edited_output['적용납품가'].to_dict()

        # 4. 분석 로직 (V2.1 계산식 절대 고정)
        COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS = \
            'SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '현재 상품수익', '현재 롤수익', '추천가 상품수익', '추천가 롤수익', '방어선'

        def calc_logic(row):
            try:
                sku_id = row['SKU ID']
                applied_p = price_map.get(sku_id, clean_num(row.get(orig_price_col, 0)))
                
                # 규격 데이터 (제시해주신 기준 수식 그대로 보존)
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 1200))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # 단위 원가 계산
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                s_r, j_r, a_r = (v/100 if v > 1 else v for v in [s_val, j_val, a_val])
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                # 제조 원가 및 추천가 산출 (비중 0.000184 고정)
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20))
                indiv_target = target_val / 100 if target_val > 1 else target_val
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                
                # 수익 분석 데이터
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                current_unit_profit = applied_p - total_box_cost
                current_roll_profit = current_unit_profit * boxes_per_roll
                rec_unit_profit = rec_nap_ga - total_box_cost
                rec_roll_profit = rec_unit_profit * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id, row.get('상품명', ''), fmt(total_box_cost), fmt(applied_p), 
                    fmt(rec_nap_ga), fmt(current_unit_profit), fmt(current_roll_profit),
                    fmt(rec_unit_profit), fmt(rec_roll_profit), status
                ], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS])
            except:
                return pd.Series([row.get('SKU ID', ''), row.get('상품명', ''), '0원', '0원', '0원', '0원', '0원', '0원', '0원', '오류'], 
                                 index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_REC, COL_UNIT_PROFIT, COL_PROFIT, COL_REC_UNIT_PROFIT, COL_REC_PROFIT, COL_STATUS])

        # 5. 리포트 출력
        df_res = df_products.apply(calc_logic, axis=1)
        st.subheader(f"📊 {selected_month} 마진 상세 분석")
        
        st.dataframe(
            df_res.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=[COL_STATUS]), 
            use_container_width=True, hide_index=True
        )

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_margin_calc()
