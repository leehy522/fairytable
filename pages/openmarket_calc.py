import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import os
import datetime # [추가] 날짜 인식을 위한 모듈

def show_openmarket_calc():
    st.title("🛍️ 오픈마켓 수익 분석 시뮬레이터 (V2.2 - 현재 월 자동 로드)")
    st.markdown("---")

    try:
        # 1. 데이터 로드
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

        df_products['SKU ID'] = df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # ---------------------------------------------------------
        # [핵심] 2. 현재 월 자동 기본값 설정 로직
        # ---------------------------------------------------------
        now = datetime.datetime.now()
        month_options = df_costs['월'].unique().tolist()
        
        # 기본 인덱스 설정 (찾지 못할 경우 0번째)
        default_month_index = 0
        
        # 시트의 '월' 형식이 '2026-04' 혹은 '4월' 등일 경우를 대비한 매칭 작업
        for i, m_val in enumerate(month_options):
            # 현재 연도와 월이 데이터에 포함되어 있는지 확인
            if str(now.year) in m_val and str(now.month).zfill(2) in m_val:
                default_month_index = i
                break
            elif f"{now.month}월" in m_val:
                default_month_index = i
                break

        selected_month = st.selectbox(
            "📅 원가 기준 월 선택", 
            month_options, 
            index=default_month_index # 현재 월을 디폴트로 설정
        )
        # ---------------------------------------------------------

        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 3. 사이드바 및 판매가 편집기
        st.sidebar.header("⚙️ 오픈마켓 설정")
        platform = st.sidebar.selectbox("플랫폼 선택", ["네이버 스마트스토어", "알리익스프레스"])
        fee_rate = 0.06 if platform == "네이버 스마트스토어" else 0.12
        shipping_fee = st.sidebar.number_input("건당 택배비 (원)", value=2400)
        packing_extra = st.sidebar.number_input("추가 부자재비 (원)", value=0)

        st.subheader(f"✍️ {platform} 판매가 설정")
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

        # 4. 분석 로직 (V3.0 롤 무게 기반)
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

        # 5. 리포트 출력
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
        st.error(f"시스템 오류: {e}")

if check_password():
    show_openmarket_calc()
