import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
from auth import check_password  # 보안 모듈 추가

# 1. 페이지 기본 설정 및 보안 체크 (반드시 최상단에 위치)
st.set_page_config(page_title="요정비닐 원가 시뮬레이터", page_icon="📊", layout="wide")

if not check_password():
    st.stop()

st.title("📊 요정비닐 원가 및 마진 시뮬레이터")
st.markdown("---")

# 2. 구글 시트 연결
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"연결 객체 생성 실패. secrets.toml 형식을 재확인하십시오: {e}")
    st.stop()

# 반드시 실제 스프레드시트 주소로 변경할 것
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/본인의_스프레드시트_ID_입력/edit"

# 3. 데이터 호출
try:
    df_products = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="상품DB", ttl=0)
    df_costs = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="월별단가", ttl=0)
    
    df_products = df_products.dropna(how="all")
    df_costs = df_costs.dropna(how="all")
except Exception as e:
    st.error(f"데이터 호출 실패. 공유 권한 설정이나 시트명('상품DB', '월별단가')이 정확한지 확인하십시오: {e}")
    st.stop()

# 4. 마진 계산 로직
def calculate_margin(products, costs, target_month):
    target_cost = costs[costs['적용월'] == target_month]
    
    if target_cost.empty:
        st.warning(f"'{target_month}' 단가 데이터가 없습니다. 시트의 가장 최근 단가를 적용합니다.")
        target_cost = costs.iloc[-1:]
        
    try:
        sinjae = float(target_cost['신재'].values[0])
        pigment = float(target_cost['안료'].values[0])
    except (IndexError, ValueError):
        st.error("단가 데이터 오류: '신재', '안료' 열에 유효한 숫자가 입력되었는지 확인하십시오.")
        return products

    st.info(f"💡 적용 원자재 단가 - 신재: {sinjae}원/kg | 안료: {pigment}원/kg")

    # 무게 및 원가 계산 (가로 x 세로 x 두께 x 0.92(비중) x 2면 / 10000)
    products['1장무게(kg)'] = (products['가로(cm)'] * products['세로(cm)'] * products['두께(T)'] * 0.92 * 2) / 10000
    products['1장원가(원)'] = products['1장무게(kg)'] * (sinjae + pigment)
    
    # 마진율 계산 = (판매가 - 원가) / 판매가 * 100
    products['마진율(%)'] = ((products['현재판매가'] - products['1장원가(원)']) / products['현재판매가']) * 100
    
    # 데이터 정리
    products['마진율(%)'] = products['마진율(%)'].round(2)
    products['1장원가(원)'] = products['1장원가(원)'].astype(int)
    
    return products

# 5. 시뮬레이션 UI
with st.expander("🛠️ 데이터 연동 및 조건 설정", expanded=True):
    col1, col2 = st.columns(2)
    
    with col1:
        month_list = df_costs['적용월'].dropna().unique().tolist()
        default_idx = len(month_list) - 1 if month_list else 0
        
        if month_list:
            selected_month = st.selectbox("적용 기준월 선택", month_list, index=default_idx)
        else:
            selected_month = st.text_input("적용월 수동 입력", value=datetime.now().strftime('%Y-%m'))
            
    with col2:
        st.write("연동 상태: 🟢 정상")
        if st.button("🔄 최신 데이터 강제 동기화"):
            st.cache_data.clear()
            st.rerun()

# 6. 결과 테이블 출력
if not df_products.empty and not df_costs.empty:
    st.subheader(f"✅ {selected_month} 기준 마진 현황")
    final_df = calculate_margin(df_products.copy(), df_costs.copy(), selected_month)
    
    st.dataframe(
        final_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "현재판매가": st.column_config.NumberColumn("현재판매가", format="%d 원"),
            "1장원가(원)": st.column_config.NumberColumn("1장원가(원)", format="%d 원"),
            "마진율(%)": st.column_config.NumberColumn("마진율(%)", format="%.2f %%"),
        }
    )
else:
    st.warning("스프레드시트에 데이터가 비어있습니다. 데이터를 입력하십시오.")
