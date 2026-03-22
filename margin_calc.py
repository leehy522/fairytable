import streamlit as st
import pandas as pd

# 페이지 기본 설정 (메뉴 아이콘과 제목)
st.set_page_config(page_title="요정비닐 원가 시뮬레이터", page_icon="📊", layout="wide")

st.title("📊 실시간 원가 및 마진 시뮬레이터")
st.markdown("---")

# 안내 메시지
st.info("💡 구글 스프레드시트의 '월별단가'와 '상품DB'를 불러와 실시간 마진율을 계산합니다.")

# 1. 시뮬레이션 컨트롤 패널 (화면 상단)
with st.expander("🛠️ 시뮬레이션 설정", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        # 나중에 구글 시트에서 월을 자동으로 불러오겠지만, 수동 선택 기능도 남겨둡니다.
        selected_month = st.selectbox("적용 기준월", ["2026-03", "2026-04", "2026-05"])
    with col2:
        st.write("현재 연동 상태: 🔴 구글 시트 미연결")
        if st.button("🔄 최신 단가 불러오기"):
            st.warning("내일 구글 시트 연동 코드가 들어가면 작동합니다!")

# 2. 결과 출력 화면 (화면 하단)
st.subheader(f"✅ {selected_month} 기준 요정비닐 상품별 마진 현황")

# (임시) 데이터가 들어갈 빈 껍데기 표를 띄워둡니다.
dummy_data = pd.DataFrame(columns=["상품명", "재질", "가로", "세로", "두께", "1장무게(kg)", "1장원가(원)", "현재판매가", "마진율(%)"])
st.dataframe(dummy_data, use_container_width=True, hide_index=True)
