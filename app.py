import streamlit as st
from auth import check_password

# 1. 페이지 설정 (무조건 1번)
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 2. 강제로 기본 메뉴 숨기기 (CSS) - 로그인 전후 모두 적용
st.markdown("""
    <style>
    /* 기본 네비게이션 숨기기 */
    section[data-testid="stSidebarNav"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

# 3. 로그인 체크
if not check_password():
    st.stop()  # 로그인 안 되면 여기서 멈춤

# ─── 로그인 성공 후 실행되는 구역 ───

# 4. 메뉴 설정 (라디오 버튼)
st.sidebar.title("🚀 요정비닐 관리자")

# 💡 파일명 앞에 숫자가 있다면 아래 경로도 '01_...', '02_...'로 똑같이 맞춰야 합니다!
MENU_map = {
    "🏷️ 요정비닐 상품 현황": "pages/product_status.py",
    "🚚 밀크런 PPT 변환": "pages/milkrun_ppt.py",
    "📦 택배 송장 변환": "pages/invoice.py",
    "🏭 원가 시뮬레이터": "pages/cost_simulator.py",
    "📈 시장 지표 분석": "pages/market_index.py",
    "🏛️ 나라장터 입찰": "pages/narajangte.py"
}

selection = st.sidebar.radio("메뉴를 선택하세요", list(MENU_OPTIONS.keys()))

# 5. 페이지 이동
if selection:
    try:
        st.switch_page(MENU_OPTIONS[selection])
    except Exception as e:
        st.error(f"파일을 찾을 수 없습니다: {MENU_OPTIONS[selection]}")
        st.info("💡 깃허브의 파일명과 코드 내 경로가 일치하는지 확인해주세요.")
