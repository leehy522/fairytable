import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime
import unicodedata

def show_margin_calc():
    # 버전 표기 업데이트
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V4.0 - 논리 통합)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 (V3.9 기능 유지)
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        s1, s2, s3 = quote("상품목록"), quote("원가기준"), quote("월별납품가")
        
        df_p = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s1}")
        df_c = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s2}")
        df_m = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s3}")
        
        # 칼럼명 정규화
        for df in [df_p, df_c, df_m]:
            df.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df.columns]

        def clean_num(v):
            if pd.isna(v) or v == '': return 0
            s = re.sub(r'[^0-9.]', '', str(v))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # SKU ID 통일
        sku_col = next((c for c in df_p.columns if 'SKU' in c.upper()), df_p.columns[0])
        df_p[sku_col] = df_p[sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        m_sku_col = next((c for c in df_m.columns if 'SKU' in c.upper()), df_m.columns[0])
        df_m[m_sku_col] = df_m[m_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 사이드바 및 월 자동 선택 (V3.9 기능 유지)
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 조정 (%)", value=0.0, step=0.5)
        
        now = datetime.datetime.now()
        months = df_c['월'].astype(str).str.strip().unique().tolist()
        d_idx = next((i for i, m in enumerate(months) if str(now.year) in m and str(now.month).zfill(2) in m), 0)
        sel_month = st.selectbox("📅 원가 기준 월 선택", months, index=d_idx)
        
        # 원료 단가 추출 (비율 제외 로직)
        c_row = df_c[df_c['월'] == sel_month].iloc[0]
        sinjae_p = clean_num(next((c_row[k] for k in c_row.index if '신재' in k and '비율' not in k), 0))
        jaesaeng_p = clean_num(next((c_row[k] for k in c_row.index if '재생' in k and '비율' not in k), 0))

        # 3. 단가 열 매핑 (default 우선 로직 유지)
        target_month_col = sel_month.replace(' ', '')
        if target_month_col in df_m.columns:
            p_col = target_month_col
        else:
            p_col = next((c for c in df_m.columns if c.lower() == 'default'), None)
            if not p_col: p_col = next((c for c in df_m.columns if '기본' in c), df_m.columns[2])
        
        st.info(f"✅ 납품가 기준: '{p_col}' 열 데이터 적용 중")
        m_price_dict = df_m.set_index(m_sku_col)[p_col].to_dict()

        # 4. 시뮬레이션 데이터 에디터 (V3.9 기능 유지)
        edit_df = df_p[[sku_col, '상품명']].copy()
        edit_df.rename(columns={sku_col: 'SKU ID'}, inplace=True)
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(m_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct/100)).round(0).astype(int)
        
        edit_df.set_index('SKU ID', inplace=True)
        e_output = st.data_editor(edit_df, use_container_width=True, key="v40_sync")
        p_map = e_output['적용납품가'].to_dict()

        # ---------------------------------------------------------
        # [신규 로직 이식] 5. 분석 로직 (V4.0 체질 개선 엔진)
        # ---------------------------------------------------------
        def calc_logic(row):
            try:
                sku = row[sku_col]
                applied_p = p_map.get(sku, 0)
                
                # 규격 파싱
                garo = clean_num(row.get('가로', 90))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # 원재료 배합비
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae_p * s_r) + (jaesaeng_p * j_r)

                # 물리적 무게 산출
                roll_weight = garo * dukki * length * 0.0184
                total_pcs = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                b_per_r = total_pcs / box_pcs if box_pcs > 0 else 1
                
                # [체질 개선 1] 불량 리스크 5% 반영
                mfg_cost_raw = ((roll_weight * unit_price) / b_per_r) + box_cost if b_per_r > 0 else 0
                mfg_cost = round(mfg_cost_raw * 1.05, 0) # 리스크 프리미엄 추가
                
                # [체질 개선 2] 추천가 논리 통합 (마진 20% vs 롤당 수익 20,000원 중 높은 값)
                min_roll_profit_goal = 20000 
                rec_by_margin = mfg_cost / 0.8
                rec_by_absolute = mfg_cost + (min_roll_profit_goal / b_per_r)
                rec_p = round(max(rec_by_margin, rec_by_absolute), 0)
                
                roll_profit = (applied_p - mfg_cost) * b_per_r
                
                # [체질 개선 3] 경고 기준값 강화 (17,000원 적자 / 20,000원 주의)
                status = "✅ 정상"
                if roll_profit < 17000: status = "🚨 적자위험"
                elif roll_profit < 20000: status = "⚠️ 수익주의"
                elif roll_profit > 60000: status = "💰 고마진"

                return pd.Series([sku, row.get('상품명',''), f"{roll_weight:.2f}kg", mfg_cost, applied_p, rec_p, (applied_p - mfg_cost), roll_profit, status],
                                 index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])
            except:
                return pd.Series([sku, '오류', '0kg', 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])

        # 6. 결과 리포트 및 다운로드 (V3.9 UI 유지 + 경고색 추가)
        df_res = df_p.apply(calc_logic, axis=1)
        st.subheader(f"📊 {sel_month} 마진 정밀 분석 결과")
        
        df_disp = df_res.copy()
        for c in ['제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익']:
            df_disp[c] = df_disp[c].apply(lambda x: f"{int(x):,}원")

        # 스타일링: 적자(빨강)와 주의(노랑) 모두 표시
        def highlight_status(v):
            if '🚨' in str(v): return 'color: red; font-weight: bold;'
            if '⚠️' in str(v): return 'color: #ffaa00; font-weight: bold;'
            return ''

        st.dataframe(df_disp.style.applymap(highlight_status, subset=['방어선']), use_container_width=True, hide_index=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='분석결과')
        st.download_button("📥 엑셀 리포트 다운로드", buffer.getvalue(), f"마진분석_V4_{sel_month}.xlsx")

    except Exception as e:
        st.error(f"⚠️ 시스템 오류: {e}")

if check_password():
    show_margin_calc()
