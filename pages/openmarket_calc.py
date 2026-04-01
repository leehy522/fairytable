import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import os

def show_openmarket_calc():
    st.title("🛍️ 오픈마켓 수익 분석 시뮬레이터 (V1.0)")
    st.markdown("---")

    try:
        # 1. 데이터 로드 (V2.1 표준 시트 동일 사용)
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

        # 2. 고정 변수 설정 (네이버, 알리, 택배비)
        st.sidebar.header("⚙️ 오픈마켓 설정")
        platform = st.sidebar.selectbox("플랫폼 선택", ["네이버 스마트스토어", "알리익스프레스"])
        
        # 플랫폼별 수수료 자동 세팅
        fee_rate = 0.06 if platform == "네이버 스마트스토어" else 0.12
        st.sidebar.write(f"현재 수수료율: {fee_rate*100:.1f}%")
        
        shipping_fee = st.sidebar.number_input("건당 택배비 (원)", value=2400)
        packing_extra = st.sidebar.number_input("추가 부자재비 (봉투/테이프 등)", value=0)

        # 3. 원가 기준 설정
        selected_month = st.selectbox("📅 원가 기준 월 선택", df_costs['월'].unique().tolist())
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 4. 판매가 시뮬레이션 표
        st.subheader(f"✍️ {platform} 판매가 설정")
        st.caption("판매가를 수정하면 수수료와 택배비를 제외한 순수익이 계산됩니다.")
        
        edit_df = df_products[['SKU ID', '상품명']].copy()
        edit_df['SKU ID'] = edit_df['SKU ID'].astype(str).str.split('.').str[0]
        # 오픈마켓은 쿠팡 납품가와 다르므로 시뮬레이션용 임시 판매가(디폴트 10,000원) 설정
        edit_df['설정판매가'] = 10000 

        edited_output = st.data_editor(
            edit_df,
            column_config={
                "SKU ID": st.column_config.TextColumn("SKU ID", disabled=True),
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "설정판매가": st.column_config.NumberColumn("판매가(수정)", format="%d원")
            },
            hide_index=True, use_container_width=True, key="openmarket_editor"
        )
        price_map = edited_output.set_index('SKU ID')['설정판매가'].to_dict()

        # 5. 수익 분석 로직 (V2.1 표준 제조원가 로직 계승)
        def calc_open_logic(row):
            try:
                sku_id = str(row.get('SKU ID', '')).split('.')[0]
                selling_p = price_map.get(sku_id, 10000)
                
                # [표준 로직] 제조원가 산출
                garo = clean_num(row.get('가로', 90))
                sero = clean_num(row.get('세로', 100))
                dukki = clean_num(row.get('두께', 0.0125))
                box_pcs = clean_num(row.get('매수', 100))
                box_cost = clean_num(row.get('박스비', 0))
                
                s_r = clean_num(next((row[k] for k in row.index if '신재' in k and '비율' in k), 100)) / 100
                j_r = clean_num(next((row[k] for k in row.index if '재생' in k and '비율' in k), 0)) / 100
                unit_price = (sinjae * s_r) + (jaesaeng * j_r)

                single_weight = garo * sero * dukki * 0.000184
                total_mfg_cost = round((single_weight * box_pcs * unit_price) + box_cost, 0)
                
                # [오픈마켓 로직] 순수익 산출
                platform_fee = selling_p * fee_rate
                total_out_cost = total_mfg_cost + platform_fee + shipping_fee + packing_extra
                net_profit = selling_p - total_out_cost
                margin_rate = (net_profit / selling_p) * 100 if selling_p > 0 else 0
                
                def fmt(v): return f"{int(round(v, 0)):,}원"

                return pd.Series([
                    sku_id, row.get('상품명', ''), fmt(total_mfg_cost), fmt(selling_p), 
                    fmt(platform_fee), fmt(shipping_fee), fmt(net_profit), f"{margin_rate:.1f}%"
                ], index=['SKU ID', '상품명', '제조원가', '판매가', '수수료', '택배비', '최종수익', '마진율'])
            except:
                return pd.Series([sku_id, '오류', '0원', '0원', '0원', '0원', '0원', '0%'], 
                                 index=['SKU ID', '상품명', '제조원가', '판매가', '수수료', '택배비', '최종수익', '마진율'])

        # 6. 결과 출력
        df_res = df_products.apply(calc_open_logic, axis=1)
        st.subheader(f"📊 {platform} 마진 분석 리포트")
        st.dataframe(df_res, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"시스템 오류: {e}")

if check_password():
    show_openmarket_calc()
