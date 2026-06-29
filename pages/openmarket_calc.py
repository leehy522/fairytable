import streamlit as st
import pandas as pd
from urllib.parse import quote
from auth import check_password
import io
import re
import os
import datetime 
import unicodedata # [추가] 구글 시트 한글 깨짐 및 자소 분리 방지용

def show_openmarket_calc():
    st.title("🛍️ 오픈마켓 수익 분석 시뮬레이터 (V2.3 - 네이버 판매가 월별 동적 연동)")
    st.markdown("---")

    try:
        # 1. 데이터 원격 로드 및 시트 정의
        sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
        sheet_name_1 = quote("상품목록")      # "상품목록" URL 인코딩
        sheet_name_2 = quote("원가기준")      # "원가기준" URL 인코딩
        sheet_name_3 = quote("네이버 판매가") # [추가] "네이버 판매가" 타겟 시트 URL 인코딩
        
        df_products = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_1}")
        df_costs = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_2}")
        df_sales = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name_3}")
        
        # [정문화] 맥(Mac) 환경 등에서 발생하는 한글 자소 분리(NFD) 현상 및 공백 일괄 청소
        for df in [df_products, df_costs, df_sales]:
            df.columns = [unicodedata.normalize('NFC', str(c)).strip() for c in df.columns]

        # 숫자 정제 헬퍼 함수
        def clean_num(value):
            if pd.isna(value) or value == '': return 0
            s = re.sub(r'[^0-9.]', '', str(value))
            return pd.to_numeric(s, errors='coerce') if s else 0

        # 데이터 키 값 및 텍스트 타입 동기화
        df_products['SKU ID'] = df_products['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df_costs['월'] = df_costs['월'].astype(str).str.strip()
        
        # '네이버 판매가' 시트의 SKU 매핑용 기준 열 자동 확보 및 정규화
        sales_sku_col = next((c for c in df_sales.columns if 'SKU' in c.upper()), df_sales.columns[0])
        df_sales[sales_sku_col] = df_sales[sales_sku_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        # 2. 현재 월 자동 기본값 설정 로직
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

        selected_month = st.selectbox(
            "📅 원가 및 판매가 기준 월 선택", 
            month_options, 
            index=default_month_index
        )

        # 3. 선택 월 기준 원가 데이터 파싱
        target_cost_row = df_costs[df_costs['월'] == selected_month].iloc[0]
        sinjae = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '신재' in k), 0))
        jaesaeng = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '재생' in k), 0))
        anlyo = clean_num(next((target_cost_row[k] for k in target_cost_row.index if '안료' in k), 0))

        # 4. [핵심] '네이버 판매가' 시트에서 선택된 월(예: 2026.06) 컬럼 동적 매핑 엔진
        target_sale_col = selected_month.replace(' ', '') # 공백 제거하여 시트의 헤더와 비교 연동
        
        if target_sale_col in df_sales.columns:
            st.success(f"📈 판매가 동기화 완료: '네이버 판매가' 시트의 **[{target_sale_col}]** 열을 정상적으로 호출했습니다.")
            sales_price_dict = df_sales.set_index(sales_sku_col)[target_sale_col].to_dict()
        else:
            # 매칭되는 월 컬럼이 존재하지 않을 때 차선책(Fallback)으로 첫 번째 금액 데이터 열 추적
            fallback_col = df_sales.columns[2] if len(df_sales.columns) > 2 else df_sales.columns[-1]
            st.warning(f"⚠️ '네이버 판매가' 시트에 일치하는 [{target_sale_col}] 열이 확인되지 않아 기본 열 [{fallback_col}]로 자동 대체합니다.")
            sales_price_dict = df_sales.set_index(sales_sku_col)[fallback_col].to_dict()

        # 5. 사이드바 설정 부문
        st.sidebar.header("⚙️ 오픈마켓 설정")
        platform = st.sidebar.selectbox("플랫폼 선택", ["네이버 스마트스토어", "알리익스프레스"])
        fee_rate = 0.06 if platform == "네이버 스마트스토어" else 0.12
        shipping_fee = st.sidebar.number_input("건당 택배비 (원)", value=2400)
        packing_extra = st.sidebar.number_input("추가 부자재비 (원)", value=0)

        st.subheader(f"✍️ {platform} 실시간 판매가 세팅 (시트 기본값 자동 반영)")
        
        # 상품목록 뼈대에 구글 시트에서 가져온 월별 판매가를 디폴트 값으로 주입
        edit_df = df_products[['SKU ID', '상품명']].copy()
        edit_df['설정판매가'] = edit_df['SKU ID'].apply(lambda x: clean_num(sales_price_dict.get(x, 0)))
        edit_df.set_index('SKU ID', inplace=True)
        
        # 데이터 에디터 출력 및 수정값 실시간 수집
        edited_output = st.data_editor(
            edit_df,
            column_config={
                "상품명": st.column_config.TextColumn("상품명", disabled=True),
                "설정판매가": st.column_config.NumberColumn("판매가(수정 가능)", format="%d원")
            },
            use_container_width=True, 
            key="openmarket_sync_editor_v5" 
        )
        price_map = edited_output['설정판매가'].to_dict()

        # 6. 정밀 마진 분석 알고리즘 (V3.0 롤 무게 상수 0.0184 고정)
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

                # 현장 실측 원단 무게 및 롤당 카운팅 기반 원가 도출
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

        # 7. 분석 대시보드 출력 리포트
        df_res = df_products.apply(calc_open_logic, axis=1)
        st.subheader(f"📊 {selected_month} 플랫폼별 최종 마진 시뮬레이션 결과")
        
        # 마이너스 수익 발생 시 붉은색 볼드체 경고 포맷팅
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
