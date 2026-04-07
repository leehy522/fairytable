import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime
import unicodedata

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V4.2)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 및 정규화
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        s1, s2, s3 = quote("상품목록"), quote("원가기준"), quote("월별납품가")
        
        df_p = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s1}")
        df_c = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s2}")
        df_m = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s3}")
        
        for df in [df_p, df_c, df_m]:
            df.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df.columns]

        def clean_num(v):
            if pd.isna(v) or v == '': return 0
            s = re.sub(r'[^0-9.]', '', str(v))
            return pd.to_numeric(s, errors='coerce') if s else 0

        sku_col = next((c for c in df_p.columns if 'SKU' in c.upper()), df_p.columns[0])
        df_p[sku_col] = df_p[sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        m_sku_col = next((c for c in df_m.columns if 'SKU' in c.upper()), df_m.columns[0])
        df_m[m_sku_col] = df_m[m_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 사이드바 및 월 선택
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 조정 (%)", value=0.0, step=0.5)
        
        now = datetime.datetime.now()
        months = df_c['월'].astype(str).str.strip().unique().tolist()
        d_idx = next((i for i, m in enumerate(months) if str(now.year) in m and str(now.month).zfill(2) in m), 0)
        sel_month = st.selectbox("📅 원가 기준 월 선택", months, index=d_idx)
        
        c_row = df_c[df_c['월'] == sel_month].iloc[0]
        sinjae_p = clean_num(next((c_row[k] for k in c_row.index if '신재' in k and '비율' not in k), 0))
        jaesaeng_p = clean_num(next((c_row[k] for k in c_row.index if '재생' in k and '비율' not in k), 0))

        # 3. 단가 열 매핑
        target_month_col = sel_month.replace(' ', '')
        if target_month_col in df_m.columns:
            p_col = target_month_col
        else:
            p_col = next((c for c in df_m.columns if c.lower() == 'default'), None)
            if not p_col: p_col = next((c for c in df_m.columns if '기본' in c), df_m.columns[2])
        
        st.info(f"✅ 납품가 기준: '{p_col}' 열 적용 중")
        m_price_dict = df_m.set_index(m_sku_col)[p_col].to_dict()

        # 4. 데이터 에디터
        edit_df = df_p[[sku_col, '상품명']].copy()
        edit_df.rename(columns={sku_col: 'SKU ID'}, inplace=True)
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(m_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct/100)).round(0).astype(int)
        edit_df.set_index('SKU ID', inplace=True)
        e_output = st.data_editor(edit_df, use_container_width=True, key="v42_sync")
        p_map = e_output['적용납품가'].to_dict()

        # 5. 분석 로직 (V4.1 리스크 및 수익 하한선 유지)
        def calc_logic(row):
            try:
                sku = row[sku_col]
                applied_p = p_map.get(sku, 0)
                garo = clean_num(row.get('가로', 90))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae_p * s_r) + (jaesaeng_p * j_r)

                roll_weight = garo * dukki * length * 0.0184
                total_pcs = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                b_per_r = total_pcs / box_pcs if box_pcs > 0 else 1
                
                # 불량 리스크 5% 반영
                mfg_cost_raw = ((roll_weight * unit_price) / b_per_r) + box_cost if b_per_r > 0 else 0
                mfg_cost = round(mfg_cost_raw * 1.05, 0)
                
                # 추천가 논리 통합 (마진 20% vs 롤당 수익 20,000원)
                min_roll_profit_goal = 20000 
                rec_p = round(max(mfg_cost / 0.8, mfg_cost + (min_roll_profit_goal / b_per_r)), 0)
                
                roll_profit = (applied_p - mfg_cost) * b_per_r
                
                status = "✅ 정상"
                if roll_profit < 17000: status = "🚨 적자위험"
                elif roll_profit < 20000: status = "⚠️ 수익주의"
                elif roll_profit > 60000: status = "💰 고마진"

                return pd.Series([sku, row.get('상품명',''), f"{roll_weight:.2f}kg", mfg_cost, applied_p, rec_p, (applied_p - mfg_cost), roll_profit, status],
                                 index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])
            except:
                return pd.Series([sku, '오류', '0kg', 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])

        # ---------------------------------------------------------
        # [신규] 6. 고마진 우선순위 정렬 및 출력
        # ---------------------------------------------------------
        df_res = df_p.apply(calc_logic, axis=1)
        
        # 롤당수익 기준으로 내림차순 정렬 (높은 수익이 위로)
        df_res = df_res.sort_values(by='롤당수익', ascending=False)
        
        st.subheader(f"📊 {sel_month} 마진 정밀 분석 (우선순위 정렬)")
        
        df_disp = df_res.copy()
        for c in ['제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익']:
            df_disp[c] = df_disp[c].apply(lambda x: f"{int(x):,}원")

        def highlight_status(v):
            if '🚨' in str(v): return 'color: red; font-weight: bold;'
            if '⚠️' in str(v): return 'color: #ffaa00; font-weight: bold;'
            if '💰' in str(v): return 'color: #00ff00; font-weight: bold;'
            return ''

        st.dataframe(df_disp.style.map(highlight_status, subset=['방어선']), use_container_width=True, hide_index=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='분석결과')
        st.download_button("📥 우선순위 리포트 다운로드", buffer.getvalue(), f"고마진_우선순위_{sel_month}.xlsx")

    except Exception as e:
        st.error(f"⚠️ 시스템 오류: {e}")

if check_password():
    show_margin_calc()
