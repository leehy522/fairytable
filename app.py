import streamlit as st
from auth import check_password

# 1. 페이지 설정 (가장 먼저)
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 2. [필살기] 스트림릿 자동 메뉴 숨기기 (CSS)
# ★ 로그인 체크보다 위에 있어야 합니다! ★
st.markdown("""
    <style>
    /* 1. 사이드바의 자동 네비게이션 숨기기 */
    [data-testid="stSidebarNav"] {display: none !important;}
    
    /* 2. 사이드바 상단 여백 제거 */
    [data-testid="stSidebarNavContent"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

# 3. 로그인 체크
if not check_password():
    st.stop()

# ─── 로그인 성공 후 실행되는 구역 ───

# 4. 메뉴 설정 (라디오 버튼)
st.sidebar.title("🚀 요정비닐 관리자")

# 💡 파일명 앞에 숫자가 있다면 아래 경로도 '01_...', '02_...'로 똑같이 맞춰야 합니다!
MENU_MAP = {
    "🏷️ 요정비닐 상품 현황": "product_status.py",
    "🚚 밀크런 PPT 변환": "milkrun_ppt.py",
    "📦 택배 송장 변환": "invoice.py",
    "🏭 원가 시뮬레이터": "cost_simulator.py",
    "📈 시장 지표 분석": "market_index.py",
    "🏛️ 나라장터 입찰": "narajangte.py"
}

selection = st.sidebar.radio("메뉴를 선택하세요", list(MENU_MAP.keys()))

# 5. 페이지 이동
if selection:
    try:
        st.switch_page(MENU_MAP[selection])
    except Exception as e:
        st.error(f"파일을 찾을 수 없습니다: {MENU_MAP[selection]}")
        st.info("💡 깃허브의 파일명과 코드 내 경로가 일치하는지 확인해주세요.")
