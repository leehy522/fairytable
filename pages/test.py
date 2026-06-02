import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import datetime
import unicodedata

# ----------------------------------------------------------------
# 1. [공통 레벨] 데이터 정규화 및 정리 유틸리티 함수
# ----------------------------------------------------------------
def clean_num(v):
    if pd.isna(v) or v == '': return 0
    s = re.sub(r'[^0-9.]', '', str(v))
    return pd.to_numeric(s, errors='coerce') if s else 0

# ----------------------------------------------------------------
# 2. [탭 1] 페어리테이블 마진 정밀 시뮬레이터 (V4.2)
# ----------------------------------------------------------------
def show_margin_calc():
    st.title("🛡️ 페어리테이블 마진 정밀 시뮬레이터 (V4.2)")
    st.markdown("---")

    try:
        # 데이터 로드 및 정규화
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        s1, s2, s3 = quote("상품목록"), quote("원가기준"), quote("월별납품가")
        
        df_p = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s1}")
        df_c = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s2}")
        df_m = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={s3}")
        
        for df in [df_p, df_c, df_m]:
            df.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df.columns]

        sku_col = next((c for c in df_p.columns if 'SKU' in c.upper()), df_p.columns[0])
        df_p[sku_col] = df_p[sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        m_sku_col = next((c for c in df_m.columns if 'SKU' in c.upper()), df_m.columns[0])
        df_m[m_sku_col] = df_m[m_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 사이드바 설정
        st.sidebar.header("⚙️ 시뮬레이션 설정")
        adj_pct = st.sidebar.number_input("📈 전 품목 일괄 조정 (%)", value=0.0, step=0.5, key="margin_adj_pct")
        
        now = datetime.datetime.now()
        months = df_c['월'].astype(str).str.strip().unique().tolist()
        d_idx = next((i for i, m in enumerate(months) if str(now.year) in m and str(now.month).zfill(2) in m), 0)
        sel_month = st.selectbox("📅 원가 기준 월 선택", months, index=d_idx, key="margin_month_select")
        
        c_row = df_c[df_c['월'] == sel_month].iloc[0]
        sinjae_p = clean_num(next((c_row[k] for k in c_row.index if '신재' in k and '비율' not in k), 0))
        jaesaeng_p = clean_num(next((c_row[k] for k in c_row.index if '재생' in k and '비율' not in k), 0))

        # 단가 열 매핑
        target_month_col = sel_month.replace(' ', '')
        if target_month_col in df_m.columns:
            p_col = target_month_col
        else:
            p_col = next((c for c in df_m.columns if c.lower() == 'default'), None)
            if not p_col: p_col = next((c for c in df_m.columns if '기본' in c), df_m.columns[2])
        
        st.info(f"✅ 납품가 기준: '{p_col}' 열 적용 중 (부가세 별도)")
        m_price_dict = df_m.set_index(m_sku_col)[p_col].to_dict()

        # 데이터 에디터
        edit_df = df_p[[sku_col, '상품명']].copy()
        edit_df.rename(columns={sku_col: 'SKU ID'}, inplace=True)
        edit_df['적용납품가'] = edit_df['SKU ID'].apply(lambda x: clean_num(m_price_dict.get(x, 0)))
        edit_df['적용납품가'] = (edit_df['적용납품가'] * (1 + adj_pct/100)).round(0).astype(int)
        edit_df.set_index('SKU ID', inplace=True)
        e_output = st.data_editor(edit_df, use_container_width=True, key="v42_sync")
        p_map = e_output['적용납품가'].to_dict()

        # 분석 로직 (실측치 표준인 0.0184로 무게 상수 동기화)
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

                # 현장 실측치 매칭 0.0184 적용
                roll_weight = garo * dukki * length * 0.0184
                total_pcs = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                b_per_r = total_pcs / box_pcs if box_pcs > 0 else 1
                
                mfg_cost_raw = ((roll_weight * unit_price) / b_per_r) + box_cost if b_per_r > 0 else 0
                mfg_cost = round(mfg_cost_raw * 1.05, 0) # 불량 리스크 5% 반영
                
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

        df_res = df_p.apply(calc_logic, axis=1)
        df_res = df_res.sort_values(by='롤당수익', ascending=False)
        
        st.subheader(f"📊 {sel_month} 마진 정밀 분석 (고마진 우선 정렬)")
        
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
        st.error(f"⚠️ 시뮬레이터 시스템 오류: {e}")

# ----------------------------------------------------------------
# 3. [탭 2] 오픈마켓 수익 분석 시뮬레이터 (V2.2)
# ----------------------------------------------------------------
def show_openmarket_calc():
    st.title("🛍️ 오픈마켓 수익 분석 시뮬레이터 (V2.2)")
    st.markdown("---")

    try:
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")
        sheet_name_2 = quote("원가기준")
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        
        df_products.columns = df_products.columns.str.strip()
        df_costs.columns = df_costs.columns.str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()

        df_products['SKU ID'] = df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 현재 월 자동 기본값 설정 로직
        now = datetime.datetime.now()
        month_options = df_costs['월'].unique().tolist()
        default_month_index = 0
        
        for i, m_val in enumerate(month_options):
            if str(now.year) in m_val and str(now.month).zfill(2) in m_val:
                default_month_index = i
                break
            elif f"{now.month}월" in m_val:
                default_month_index = i
                break

        selected_month = st.selectbox("📅 원가 기준 월 선택", month_options, index=default_month_index, key="open_month_select")

        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 오픈마켓 설정 사이드바
        st.sidebar.header("⚙️ 오픈마켓 설정")
        platform = st.sidebar.selectbox("플랫폼 선택", ["네이버 스마트스토어", "알리익스프레스"])
        fee_rate = 0.06 if platform == "네이버 스마트스토어" else 0.12
        shipping_fee = st.sidebar.number_input("건당 택배비 (원)", value=2400, key="open_shipping")
        packing_extra = st.sidebar.number_input("추가 부자재비 (원)", value=0, key="open_packing")

        st.subheader(f"✍️ {platform} 판매가 설정 (부가세 별도 기준 매칭)")
        open_price_col = next((k for k in df_products.columns if '오픈마켓' in k and '판매가' in k), None)
        
        edit_df = df_products[['SKU ID', '상품명']].copy()
        if open_price_col:
            edit_df['설정판매가'] = df_products[open_price_col].apply(clean_num)
        else:
            edit_df['설정판매가'] = 0
        
        edit_df.set_index('SKU ID', inplace=True)
        edited_output = st.data_editor(
            edit_df,
            column_config={
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "설정판매가": st.column_config.NumberColumn("판매가(수정)", format="%d원")
            },
            use_container_width=True, 
            key="openmarket_sync_editor_v4" 
        )
        price_map = edited_output['설정판매가'].to_dict()

        # 분석 로직 (실측치 표준인 0.0184로 동기화 완료)
        def calc_open_logic(row):
            try:
                sku_id = row['SKU ID']
                selling_p = price_map.get(sku_id, 0)
                
                garo = clean_num(row.get('가로', 90))
                dukki = clean_num(row.get('두께', 0.0125))
                length = clean_num(next((row[k] for k in row.index if '원단' in k and '길이' in k), 0))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_val = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100))
                j_val = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0))
                a_val = clean_num(next((row[k] for k in row.index if '안료' in k and '비율' in k), 0))
                s_r, j_r, a_r = (v/100 if v > 1 else v for v in [s_val, j_val, a_val])
                unit_price = (sinjae * s_r) + (jaesaeng * j_r) + (anlyo * a_r)

                # 일관된 데이터 관리를 위해 0.0184로 일괄 수정
                roll_weight = garo * dukki * length * 0.0184
                total_pcs_in_roll = clean_num(next((row[k] for k in row.index if '롤당' in k and ('수량' in k or '카운팅' in k)), 1))
                boxes_per_roll = total_pcs_in_roll / box_pcs if box_pcs > 0 else 1
                
                total_mfg_cost = round(((roll_weight * unit_price) / boxes_per_roll) + box_cost, 0) if boxes_per_roll > 0 else 0
                
                platform_fee = selling_p * fee_rate
                total_out_cost = total_mfg_cost + platform_fee + shipping_fee + packing_extra
                net_profit = selling_p - total_out_cost
                margin_rate = (net_profit / selling_p) * 100 if selling_p > 0 else 0

                def fmt(v): return f"{int(round(v, 0)):,}원"
                def fmt_wt(v): return f"{v:.2f}kg"

                return pd.Series([
                    sku_id, row.get('상품명', ''), fmt_wt(roll_weight), fmt(total_mfg_cost), fmt(selling_p), 
                    fmt(platform_fee), fmt(shipping_fee), fmt(net_profit), f"{margin_rate:.1f}%"
                ], index=['SKU ID', '상품명', '롤무게(kg)', '제조원가', '판매가', '수수료', '택배비', '최종수익', '마진율'])
            except:
                return pd.Series([row.get('SKU ID', ''), '오류', '0.00kg', '0원', '0원', '0원', '0원', '0원', '0%'], 
                                 index=['SKU ID', '상품명', '롤무게(kg)', '제조원가', '판매가', '수수료', '택배비', '최종수익', '마진율'])

        df_res = df_products.apply(calc_open_logic, axis=1)
        st.subheader(f"📊 {selected_month} 수익 시뮬레이션 결과")
        
        def highlight_loss(val):
            if isinstance(val, str) and '-' in val and ('원' in val or '%' in val):
                return 'color: #d32f2f; font-weight: bold;'
            return ''

        st.dataframe(
            df_res.style.map(highlight_loss, subset=['최종수익', '마진율']), 
            use_container_width=True, hide_index=True
        )

    except Exception as e:
        st.error(f"⚠️ 오픈마켓 시스템 오류: {e}")

# ----------------------------------------------------------------
# 4. [탭 3] 리워드 마케팅 자동화 제어기 (신규 추가)
# ----------------------------------------------------------------
def show_reward_marketing():
    st.title("🎯 페어리테이블 리워드 마케팅 자동화 제어기 (V1.0)")
    st.markdown("---")

    # 고마진 현물 리워드 원가 기준 세팅
    st.sidebar.header("🚀 마케팅 비용 설정")
    reward_mfg_cost = st.sidebar.number_input("📦 리워드 상품 제조원가 (원)", value=3500, step=100, key="reward_mfg")
    reward_shipping = st.sidebar.number_input("🚚 리워드 발송 택배비 (원)", value=2400, step=100, key="reward_ship")
    total_reward_cost = reward_mfg_cost + reward_shipping
    
    st.sidebar.info(f"💡 리뷰 1건당 실질 마케팅 지출: {total_reward_cost:,}원")

    st.subheader("📥 실시간 서포터즈 신청 현황")
    st.markdown("구글 폼이나 외부 수집 데이터를 통해 가상 유입된 내역을 자동 필터링하는 파이프라인입니다.")

    # 데모용 데이터프레임 구성
    demo_data = {
        "신청일시": ["2026-06-01 14:22", "2026-06-01 16:05", "2026-06-02 09:11", "2026-06-02 11:30"],
        "플랫폼": ["네이버 스마트스토어", "쿠팡", "네이버 스마트스토어", "네이버 스마트스토어"],
        "구매자명": ["김지은", "박상민", "이민지", "김지은"],
        "주문번호": ["20260601-9982312", "402-9918231-11203", "20260602-1102931", "20260601-9982312"],
        "포토리뷰링크": ["https://smartstore...", "https://coupang...", "https://smartstore...", "https://smartstore..."],
        "검증상태": ["대기", "대기", "대기", "대기"]
    }
    df_reward = pd.DataFrame(demo_data)

    st.markdown("#### 🛡️ 알고리즘 자동 필터링 결과")
    
    verified_rows = []
    order_counts = df_reward["주문번호"].value_counts()
    
    for idx, row in df_reward.iterrows():
        status = "✅ 검증완료"
        reason = "정상 신청"
        
        if order_counts[row["주문번호"]] > 1:
            status = "🚨 중복제외"
            reason = "동일한 주문번호로 다중 신청 감지 (체리피커)"
        elif row["플랫폼"] == "쿠팡" and not re.match(r'^\d{3}-\d{7}-\d{7}$', str(row["주문번호"])):
            status = "⚠️ 양식오류"
            reason = "쿠팡 주문번호 대시(-) 포함 자릿수 오류"
        elif row["플랫폼"] == "네이버 스마트스토어" and len(str(row["주문번호"])) < 10:
            status = "⚠️ 양식오류"
            reason = "네이버 주문번호 식별 부적합"

        row["검증상태"] = f"{status} ({reason})"
        verified_rows.append(row)
        
    df_verified = pd.DataFrame(verified_rows)

    def highlight_reward(val):
        if "✅" in str(val): return 'background-color: #e8f5e9; color: #2e7d32; font-weight: bold;'
        if "🚨" in str(val): return 'background-color: #ffebee; color: #c62828; font-weight: bold;'
        if "⚠️" in str(val): return 'background-color: #fff8e1; color: #f57f17; font-weight: bold;'
        return ''

    st.dataframe(
        df_verified.style.map(highlight_reward, subset=["검증상태"]),
        use_container_width=True, hide_index=True
    )

    st.markdown("---")
    st.subheader("🚚 리워드 상품 발송 처리")
    
    df_shipping = df_verified[df_verified["검증상태"].str.contains("✅")].copy()
    df_shipping["발송상품"] = "40L 양면 비닐봉투 1세트(리워드)"
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="금일 발송 예정 리워드", value=f"{len(df_shipping)} 건")
    with col2:
        st.metric(label="예상 마케팅 소요 비용", value=f"{len(df_shipping) * total_reward_cost:,} 원")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_shipping.to_excel(writer, index=False, sheet_name='리워드발송명단')
        
    st.download_button(
        label="📥 발송 명단 엑셀 다운로드 (우체국/택배사 업로드용)",
        data=buffer.getvalue(),
        file_name=f"페어리테이블_리워드발송_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.ms-excel"
    )

# ----------------------------------------------------------------
# 5. [메인 커널] 인증 및 탭 라우팅 엔트리포인트
# ----------------------------------------------------------------
if check_password():
    # 상단 고정형 탭 배치로 화면 레이아웃 최적화
    main_tabs = st.tabs(["🛡️ 마진 정밀 시뮬레이터", "🛍️ 오픈마켓 수익 분석", "🎯 리워드 마케팅 제어"])
    
    with main_tabs[0]:
        show_margin_calc()
        
    with main_tabs[1]:
        show_openmarket_calc()
        
    with main_tabs[2]:
        show_reward_marketing()
