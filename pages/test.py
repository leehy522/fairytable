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
        # 1. 구글 시트 데이터 로드
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

        # 2. 원가 기준월 및 단가 설정
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        # [데이터 체크] 2월 기준 1,520원 등 원료가 자동 반영됨
        sinjae = clean_num(target_cost_row.get('신재', 0))
        jaesaeng = clean_num(target_cost_row.get('재생', 0))
        anlyo = clean_num(target_cost_row.get('안료', 0))

        # ---------------------------------------------------------
        # [핵심] 3. 납품가 편집기 (실시간 반영의 엔진)
        # ---------------------------------------------------------
        st.subheader("✍️ 납품 단가 시뮬레이션 (수정 즉시 하단 결과 반영)")
        
        orig_price_col = next((k for k in df_products.columns if '납품가' in k), None)
        edit_df = df_products[['SKU ID', '상품명']].copy()
        edit_df['적용납품가'] = df_products[orig_price_col].apply(clean_num)

        # 데이터 에디터: key를 지정하여 수정 즉시 페이지 전체가 Rerun되도록 함
        edited_output = st.data_editor(
            edit_df,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "적용납품가": st.column_config.NumberColumn("납품가(직접수정)", format="%d원")
            },
            hide_index=True,
            use_container_width=True,
            key="realtime_editor"
        )
        
        # 수정된 단가를 SKU별로 매핑
        price_map = edited_output.set_index('SKU ID')['적용납품가'].to_dict()

        # ---------------------------------------------------------
        # 4. 분석 리포트 생성 로직
        # ---------------------------------------------------------
        COL_SKU = 'SKU ID'
        COL_NAME = '상품명'
        COL_COST = '제조원가(박스)'
        COL_APPLIED = '적용납품가'
        COL_RECOMMEND = '추천납품가'
        COL_ADJ = '단가 조정액(+/-)'
        COL_COUPANG = '쿠팡판매가(42%)'
        COL_PROFIT = '롤당수익'
        COL_STATUS = '방어선'

        def calc_logic(row):
            try:
                sku_id = str(row.get('SKU ID', ''))
                # [수정 포인트] 에디터에서 수정한 값을 최우선으로 가져옴
                applied_p = price_map.get(sku_id, clean_num(row.get(orig_price_col, 0)))
                
                # 규격 데이터 정밀 파싱 (60L: 63x90x0.009 등)
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 0))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 400))
                
                # 배합비 및 단위 원가 계산
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                a_r = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                # [원가 계산 공식] (가로 * 세로 * 두께 * 비중) * 매수 * 원료가 + 박스비
                # 세로가 없는 경우 원단 길이를 기준으로 계산하는 예외 처리 포함
                effective_sero = sero if sero > 0 else (length / box_pcs if box_pcs > 0 else 100)
                single_weight = garo * effective_sero * dukki * 0.000184
                total_box_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                # [추천가 및 조정액] 목표 수익률 20% 기준 (시트의 목표 수익률 컬럼 참조)
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                target_margin = clean_num(row.get(target_col, 20)) / 100
                rec_nap_ga = round(total_box_cost / (1 - target_margin), 0)
                adj_val = rec_nap_ga - applied_p

                # [쿠팡 및 수익]
                coupang_selling = round(applied_p / 0.58, 0)
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                current_roll_profit = (applied_p - total_box_cost) * boxes_per_roll
                
                status = "✅ 정상"
                if current_roll_profit < 15000: status = "🚨 적자위험"
                elif current_roll_profit > 50000: status = "⚠️ 고마진"

                def fmt(v): return f"{int(round(v, 0)):,}원"
                def fmt_adj(v): return f"{'+' if v > 0 else ''}{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id.split('.')[0],
                    row.get('상품명', ''),
                    fmt(total_box_cost),
                    fmt(applied_p),
                    fmt(rec_nap_ga),
                    fmt_adj(adj_val),
                    fmt(coupang_selling),
                    fmt(current_roll_profit),
                    status
                ], index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_RECOMMEND, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])
            except:
                return pd.Series([row.get('SKU ID', ''), row.get('상품명', ''), '0원', '0원', '0원', '0원', '0원', '0원', '오류'], 
                                 index=[COL_SKU, COL_NAME, COL_COST, COL_APPLIED, COL_RECOMMEND, COL_ADJ, COL_COUPANG, COL_PROFIT, COL_STATUS])

        # 5. 결과 리포트 출력
        df_res = df_products.apply(calc_logic, axis=1)
        
        st.subheader(f"📊 {selected_month} 통합 분석 리포트")
        
        def style_row(val):
            if '🚨' in str(val): return 'background-color: #ffebee; color: #c62828; font-weight: bold'
            if '⚠️' in str(val): return 'background-color: #fffde7; color: #f57f17'
            return ''

        st.dataframe(
            df_res.style.applymap(style_row, subset=[COL_STATUS]),
            use_container_width=True,
            hide_index=True
        )

        # 엑셀 다운로드 (수정된 시뮬레이션 결과 반영)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='통합분석')
        st.download_button("📥 시뮬레이션 결과 엑셀 다운로드", buffer.getvalue(), f"페어리테이블_분석_{selected_month}.xlsx")

    except Exception as e:
        st.error(f"시스템 실행 오류: {e}")

if check_password():
    show_margin_calc()
