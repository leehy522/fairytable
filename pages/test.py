import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V3.3)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 (월별납품가 탭 포함)
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        s1, s2, s3 = quote("상품목록"), quote("원가기준"), quote("월별납품가")
        
        df_p = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s1}")
        df_c = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s2}")
        df_m = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s3}")
        
        for df in [df_p, df_c, df_m]: df.columns = df.columns.str.strip()

        def clean_num(v):
            if pd.isna(v) or v == '': return 0
            s = re.sub(r'[^0-9.]', '', str(v))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # SKU ID 통일 (동기화 락)
        sku_col = next((c for c in df_p.columns if 'SKU' in c.upper()), df_p.columns[0])
        df_p[sku_col] = df_p[sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        m_sku_col = next((c for c in df_m.columns if 'SKU' in c.upper()), df_m.columns[0])
        df_m[m_sku_col] = df_m[m_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 사이드바 - 일괄 조정 및 현재 월 자동 로드
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 가격 조정 (%)", value=0.0, step=0.5)
        
        now = datetime.datetime.now()
        months = df_c['월'].astype(str).str.strip().unique().tolist()
        d_idx = next((i for i, m in enumerate(months) if str(now.year) in m and str(now.month).zfill(2) in m), 0)
        sel_month = st.selectbox("📅 원가 기준 월 선택", months, index=d_idx)
        
        # 원가 데이터 추출 (비율/단가 컬럼 오인 방지)
        c_row = df_c[df_c['월'] == sel_month].iloc[0]
        sinjae_p = clean_num(next((c_row[k] for k in c_row.index if '신재' in k and '비율' not in k), 0))
        jaesaeng_p = clean_num(next((c_row[k] for k in c_row.index if '재생' in k and '비율' not in k), 0))
        anlyo_p = clean_num(next((c_row[k] for k in c_row.index if '안료' in k and '비율' not in k), 0))

        # 3. 월별 납품가 매핑
        p_col = sel_month if sel_month in df_m.columns else '기본'
        m_price_dict = df_m.set_index(m_sku_col)[p_col].to_dict()

        st.subheader("✍️ 납품 단가 시뮬레이션")
        st.info(f"✅ 현재 '{p_col}' 단가 적용 중 (일괄조정 {adj_pct}% 포함)")
        
        edit_df = df_p[[sku_col, '상품명']].copy()
        edit_df.rename(columns={sku_col: 'SKU ID'}, inplace=True)
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(m_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct/100)).round(0).astype(int)
        edit_df.set_index('SKU ID', inplace=True)
        
        e_output = st.data_editor(edit_df, use_container_width=True, key="v33_sync")
        p_map = e_output['적용납품가'].to_dict()

        # 4. 분석 로직 (롤 무게 기반 - 제조원가 증발 방지)
        def calc_logic(row):
            try:
                sku = row[sku_col]
                applied_p = p_map.get(sku, 0)
                
                # 규격 데이터
                garo = clean_num(row.get('가로', 90))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                # 배합비 및 단가 (V2.1 기준 고정)
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae_p * s_r) + (jaesaeng_p * j_r)

                # [표준 수식] 롤 무게 기반 역산 (세로값이 0이어도 원가 보존)
                roll_weight = garo * dukki * length * 0.0184
                total_pcs = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                b_per_r = total_pcs / box_pcs if box_pcs > 0 else 1
                
                mfg_cost = round(((roll_weight * unit_price) / b_per_r) + box_cost, 0) if b_per_r > 0 else 0
                
                target_col = next((k for k in row.index if '목표' in k and ('률' in k or '율' in k)), None)
                indiv_target = clean_num(row.get(target_col, 20)) / 100
                rec_p = round(mfg_cost / (1 - indiv_target), 0)
                
                roll_profit = (applied_p - mfg_cost) * b_per_r
                status = "✅ 정상"
                if roll_profit < 15000: status = "🚨 적자위험"
                elif roll_profit > 50000: status = "⚠️ 고마진"

                return pd.Series([sku, row.get('상품명',''), f"{roll_weight:.2f}kg", mfg_cost, applied_p, rec_p, (applied_p - mfg_cost), roll_profit, status],
                                 index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])
            except:
                return pd.Series([sku, '오류', '0kg', 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])

        # 5. 리포트 출력 및 엑셀 다운로드
        df_res = df_p.apply(calc_logic, axis=1)
        st.subheader(f"📊 {sel_month} 마진 분석 결과")
        
        df_disp = df_res.copy()
        for c in ['제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익']:
            df_disp[c] = df_disp[c].apply(lambda x: f"{int(x):,}원")

        st.dataframe(df_disp.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=['방어선']), use_container_width=True, hide_index=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='마진리포트')
        st.download_button("📥 엑셀 리포트 다운로드", buffer.getvalue(), f"페어리_마진분석_{sel_month}.xlsx")

    except Exception as e:
        st.error(f"⚠️ 시스템 오류: {e}")

if check_password():
    show_margin_calc()
