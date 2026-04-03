import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime

def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V3.6)")
    st.markdown("---")

    try:
        # 1. 데이터 로드
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        s1, s2, s3 = quote("상품목록"), quote("원가기준"), quote("월별납품가")
        
        df_p = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s1}")
        df_c = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s2}")
        df_m = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s3}")
        
        # 칼럼명 공백 제거
        for df in [df_p, df_c, df_m]:
            df.columns = df.columns.str.strip()

        def clean_num(v):
            if pd.isna(v) or v == '': return 0
            s = re.sub(r'[^0-9.]', '', str(v))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # SKU ID 통일 (동기화 락)
        sku_col = next((c for c in df_p.columns if 'SKU' in c.upper()), df_p.columns[0])
        df_p[sku_col] = df_p[sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        m_sku_col = next((c for c in df_m.columns if 'SKU' in c.upper()), df_m.columns[0])
        df_m[m_sku_col] = df_m[m_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 사이드바 - 설정
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 조정 (%)", value=0.0, step=0.5)
        
        now = datetime.datetime.now()
        months = df_c['월'].astype(str).str.strip().unique().tolist()
        d_idx = next((i for i, m in enumerate(months) if str(now.year) in m and str(now.month).zfill(2) in m), 0)
        sel_month = st.selectbox("📅 원가 기준 월 선택", months, index=d_idx)
        
        # 원료 단가 추출 (비율/단가 컬럼 엄격 구분)
        c_row = df_c[df_c['월'] == sel_month].iloc[0]
        sinjae_p = clean_num(next((c_row[k] for k in c_row.index if '신재' in k and '비율' not in k), 0))
        jaesaeng_p = clean_num(next((c_row[k] for k in c_row.index if '재생' in k and '비율' not in k), 0))

        # ---------------------------------------------------------
        # [수정] 3. '기본' 단가 우선 로드 로직
        # ---------------------------------------------------------
        # 1순위: 선택한 월(예: 2026-04)과 정확히 일치하는 열이 있는지 확인
        if sel_month in df_m.columns:
            p_col = sel_month
            st.info(f"✅ 특수 단가 적용: '{p_col}' 열")
        # 2순위: 일치하는 월이 없으면 '기본' 열을 사용
        elif '기본' in df_m.columns:
            p_col = '기본'
            st.info(f"✅ 기본 단가 적용: '{p_col}' 열")
        # 3순위: '기본'이라는 글자가 포함된 열이라도 찾음 (공백 등 대비)
        else:
            p_col = next((c for c in df_m.columns if '기본' in c), None)
            if not p_col:
                st.error("❌ '월별납품가' 탭에서 '기본' 열을 찾을 수 없습니다. 시트의 열 이름을 확인해주세요.")
                return
            st.warning(f"⚠️ '기본' 열을 찾지 못해 유사한 이름의 '{p_col}' 열을 적용합니다.")

        master_price_dict = df_m.set_index(m_sku_col)[p_col].to_dict()
        # ---------------------------------------------------------

        st.subheader("✍️ 납품 단가 시뮬레이션")
        edit_df = df_p[[sku_col, '상품명']].copy()
        edit_df.rename(columns={sku_col: 'SKU ID'}, inplace=True)
        
        # 단가 로드 및 일괄 조정 반영
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(master_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct / 100)).round(0).astype(int)
        
        edit_df.set_index('SKU ID', inplace=True)
        e_output = st.data_editor(edit_df, use_container_width=True, key="v36_stable")
        p_map = e_output['적용납품가'].to_dict()

        # 4. 분석 로직 (롤 무게 기반 - 제조원가 보호)
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

                # 공식: RollWeight = Garo(cm) * Dukki(mm) * Length(m) * 0.0184
                roll_weight = garo * dukki * length * 0.0184
                total_pcs = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                b_per_r = total_pcs / box_pcs if box_pcs > 0 else 1
                
                # 제조원가(박스) 역산
                mfg_cost = round(((roll_weight * unit_price) / b_per_r) + box_cost, 0) if b_per_r > 0 else 0
                
                # 목표 마진 20% 기준 추천가
                rec_p = round(mfg_cost / 0.8, 0)
                roll_profit = (applied_p - mfg_cost) * b_per_r
                
                status = "✅ 정상"
                if roll_profit < 15000: status = "🚨 적자위험"
                elif roll_profit > 50000: status = "⚠️ 고마진"

                return pd.Series([sku, row.get('상품명',''), f"{roll_weight:.2f}kg", mfg_cost, applied_p, rec_p, (applied_p - mfg_cost), roll_profit, status],
                                 index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])
            except:
                return pd.Series([sku, '오류', '0kg', 0, 0, 0, 0, 0, '오류'], index=['SKU ID', '상품명', '롤무게', '제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익', '방어선'])

        # 5. 출력 및 다운로드
        df_res = df_p.apply(calc_logic, axis=1)
        st.subheader(f"📊 {sel_month} 마진 분석 리포트")
        
        df_disp = df_res.copy()
        for c in ['제조원가', '적용납품가', '추천납품가', '상품당수익', '롤당수익']:
            df_disp[c] = df_disp[c].apply(lambda x: f"{int(x):,}원")

        st.dataframe(df_disp.style.map(lambda v: 'color: red; font-weight: bold;' if '🚨' in str(v) else '', subset=['방어선']), use_container_width=True, hide_index=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='분석결과')
        st.download_button("📥 엑셀 리포트 다운로드", buffer.getvalue(), f"페어리_마진분석_{sel_month}.xlsx")

    except Exception as e:
        st.error(f"⚠️ 시스템 오류: {e}")

if check_password():
    show_margin_calc()
