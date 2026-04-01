import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io

def show_margin_calc():
    st.title("🎯 페어리테이블 전략적 마진 시뮬레이션")
    st.markdown("---")
import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터")
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

        # 2. 원가 기준 선택
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # [오류 수정] 컬럼명에 '신재', '재생', '안료'가 포함되어 있으면 값을 가져오도록 유연하게 수정
        # 이 부분이 0이면 아래 계산식에서 비닐값이 증발합니다.
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 납품가 수동 편집 섹션
        st.subheader("✍️ 납품 단가 시뮬레이션")
        
        edit_base = df_products[['SKU ID', '상품명']].copy()
        original_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_base['수정납품가'] = df_products[original_price_col].apply(clean_num)

        # [실시간 반영용 key 추가] 
        edited_price_df = st.data_editor(
            edit_base,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "수정납품가": st.column_config.NumberColumn("납품가(수정가능)", format="%d원")
            },
            hide_index=True,
            use_container_width=True,
            key="margin_editor" 
        )
        
        price_map = edited_price_df.set_index('SKU ID')['수정납품가'].to_dict()

        # 4. 분석 로직 정의 (윤겸님의 계산식 절대 보존)
        COL_SKU, COL_NAME, COL_BOX_COST, COL_CUR_PRICE, COL_REC_PRICE, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS = \
            'SKU ID', '상품명', '제조원가(박스)', '적용납품가', '추천납품가', '단가 조정액(+/-)', '쿠팡판매가(42%)', '롤당수익', '방어선'

        def calc_logic(row):
            try:
                # [데이터 파싱]
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                # 리스트 컴프리헨션 오류 방지를 위한 안전한 파싱
                length_val = next((row[k] for k in row.index if '원단' in k and '길이' in k), 1200)
                length = clean_num(length_val)
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # [수정된 납품가 적용]
                sku_id = str(row.get('SKU ID', ''))
                applied_nap_ga = price_map.get(sku_id, clean_num(row.get(original_price_col, 0)))

                # [배합 비율 계산]
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                s_r, j_r, a_r = (v/100 if v > 1 else v for v in [s_val, j_val, a_val])
                
                # [원가 및 추천가 계산식 - 보존]
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)
                single_weight = garo * sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_val = clean_num(row.get(target_col, 20))
                indiv_target = target_val / 100 if target_val > 1 else target_val
                
                rec_nap_ga = round(total_box_cost / (1 - indiv_target), 0)
                
                # [지표 산출]
                adjustment_val = rec_nap_ga - applied_nap_ga
                coupang_selling_price = round(applied_nap_ga / 0.58, 0)

                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                current_roll_profit = (applied_nap_ga - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"
                def fmt_adj(v): return f"{'+' if v > 0 else ''}{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id.split('.')[0], row.get('상품명', ''), fmt(total_box_cost), fmt(applied_nap_ga), 
                    fmt(rec_nap_ga), fmt_adj(adjustment_val), fmt(coupang_selling_price), fmt(current_roll_profit), status
                ], index=[COL_SKU, COL_NAME, COL_BOX_COST, COL_CUR_PRICE, COL_REC_PRICE, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])
            except:
                return pd.Series(['', row.get('상품명', ''), '0원', '0원', '0원', '0원', '0원', '0원', '오류'], 
                                 index=[COL_SKU, COL_NAME, COL_BOX_COST, COL_CUR_PRICE, COL_REC_PRICE, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])

        # 결과 출력
        df_res = df_products.apply(calc_logic, axis=1)
        
        st.subheader(f"📊 실시간 분석 리포트 ({selected_month} 기준)")
        
        # [버전 호환성 수정] applymap 대신 map 사용 (최신 Pandas 에러 방지)
        st.dataframe(df_res.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=[COL_STATUS]), use_container_width=True)

    except Exception as e:
        st.error(f"실행 오류: {e}")

if check_password():
    show_margin_calc()

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        url_products = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}"
        url_costs = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}"
        
        df_products = pd.read_csv(url_products)
        df_costs = pd.read_csv(url_costs)
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = str(value).replace(',', '').replace('%', '').replace('원', '').strip()
            return pd.to_numeric(s, errors='coerce')

        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo_price = clean_num(target_cost_row.get('안료', 0))

        def calc_logic(row):
            try:
                sku_val = str(row.get('SKU ID', '')).split('.')[0]
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                indiv_target = clean_num(row.get(target_col, 20))
                if indiv_target > 1: indiv_target /= 100 

                garo = clean_num(row.get('가로', 0))
                sero = clean_num(row.get('세로', 0))
                dukki = clean_num(row.get('두께', 0))
                box_pcs = clean_num(row.get('매수', 100)) 
                box_cost = clean_num(row.get('박스비', 0))

                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                
                s_ratio = s_val / 100 if s_val > 1 else s_val
                j_ratio = j_val / 100 if j_val > 1 else j_val
                a_ratio = a_val / 100 if a_val > 1 else a_val

                # ① 비닐 1장의 무게 (kg)
                single_weight = garo * sero * dukki * 0.000184
                
                # ② 1kg당 평균 재료 단가
                unit_price = (sinjae * s_ratio) + (jaesaeng * j_ratio) + (anlyo_price * a_ratio)
                
                # ③ 최종 금액 계산 (소수점 제거 및 '원' 추가 포맷팅)
                total_cost_val = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                rec_nap_ga_val = round(total_cost_val / (1 - indiv_target), 0)
                cur_nap_ga_val = clean_num(next((row[k] for k in row.index if '납품가' in k), 0))
                adjustment_val = rec_nap_ga_val - cur_nap_ga_val
                
                # [표기용 포맷팅] 천단위 콤마와 '원' 추가
                def fmt(v): return f"{int(v):,}원"

                return pd.Series([
                    sku_val, 
                    row.get('상품명', ''), 
                    f"{int(indiv_target*100)}%", 
                    fmt(total_cost_val), 
                    fmt(cur_nap_ga_val), 
                    fmt(rec_nap_ga_val), 
                    fmt(adjustment_val)
                ])
            except Exception as e:
                return pd.Series(['', row.get('상품명', ''), '20%', '0원', '0원', '0원', '0원'])
                
        res_cols = ['SKU ID', '상품명', '설정 수익률', '제조 원가', '현재 납품가', '추천 납품가', '조정 필요액']
        df_res = df_products.apply(calc_logic, axis=1)
        df_products[res_cols] = df_res

        display_df = df_products[res_cols].copy()
        
        st.subheader(f"📊 {selected_month} 상품별 개별 마진 분석(부가세 별도)")
        
        def color_adj(val):
            if isinstance(val, (int, float)):
                return f'color: {"red" if val > 0 else "blue"}'
            return ''

        st.dataframe(display_df.style.applymap(color_adj, subset=['조정 필요액']), use_container_width=True)

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            display_df.to_excel(writer, index=False, sheet_name='상품별마진분석')
        
        st.download_button(label="📥 분석 결과 엑셀 다운로드", data=buffer.getvalue(), 
                           file_name=f"페어리테이블_마진분석_{selected_month}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"오류 발생: {e}")

if check_password():
    show_margin_calc()
