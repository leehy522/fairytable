import streamlit as st
import pandas as pd

def calculate_margin(products, costs, selected_month):
    # 1. 컬럼명 공백 제거 (KeyError 방지 핵심 논리)
    products.columns = products.columns.str.strip()
    costs.columns = costs.columns.str.strip()
    
    # 해당 월의 원가 데이터 추출
    target_cost = costs[costs['월'] == selected_month]
    if target_cost.empty:
        st.error(f"{selected_month}의 원가 데이터가 없습니다.")
        return pd.DataFrame()

    # 원가 요소 할당 (시트의 컬럼명과 정확히 일치해야 함)
    sinjae = float(target_cost['신재'].values[0])
    jaesaeng = float(target_cost['재생'].values[0])
    im_sinjae = float(target_cost['임가공(신재)'].values[0])
    im_jaesaeng = float(target_cost['임가공(재생)'].values[0])

    # 2. 계산 로직
    # 1장 원가 계산 (기존 로직 유지)
    products['1장원가(원)'] = products.apply(
        lambda row: (row['신재비율'] * sinjae + row['재생비율'] * jaesaeng + 
                     (im_sinjae if row['신재비율'] > 0 else im_jaesaeng)) * (row['가로'] * row['세로'] * row['두께'] * 0.00000184), axis=1
    )

    # 요청하신 4가지 항목 계산
    # A. 원가 (1장 원가 * 매수)
    products['원가(1장*매수)'] = (products['1장원가(원)'] * products['매수']).round(0)
    
    # B. 수익 (납품가 - 원가)
    # 시트 컬럼명이 '쿠팡 로켓 납품가(부가세 별도)'임을 가정합니다.
    products['수익'] = (products['쿠팡 로켓 납품가(부가세 별도)'] - products['원가(1장*매수)']).round(0)

    # 3. 최종 출력 항목 필터링
    display_columns = [
        '상품명', # 식별을 위해 추가
        '원가(1장*매수)', 
        '쿠팡 로켓 납품가(부가세 별도)', 
        '쿠팡 판매가', 
        '수익'
    ]
    
    return products[display_columns]

# --- 메인 실행부 ---
# 데이터프레임을 불러온 후 아래와 같이 호출하십시오.
if not df_products.empty and not df_costs.empty:
    final_df = calculate_margin(df_products.copy(), df_costs.copy(), selected_month)
    
    if not final_df.empty:
        st.subheader(f"📊 {selected_month} 마진 분석 결과")
        st.dataframe(final_df, use_container_width=True)
