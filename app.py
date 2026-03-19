import streamlit as st
from auth import check_password

# 1. 페이지 설정 (가장 먼저!)
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 2. 기본 메뉴 숨기기 (이 위치가 핵심입니다!)
st.markdown("""
    <style>
    /* 왼쪽 사이드바의 기본 네비게이션 숨기기 */
    [data-testid="stSidebarNav"] {display: none;}
    </style>
""", unsafe_allow_html=True)

# 3. 로그인 체크
if not check_password():
    st.stop()

# 4. 로그인 성공 후 나타날 윤겸님만의 메뉴 (이전 코드 계속...)
st.sidebar.title("🚀 요정비닐 관리자")
# ... 라디오 버튼 코드 ...

# 메뉴 이름과 실제 파일 경로 매칭 (파일명 앞에 숫자가 있어도 상관없음)
MENU_OPTIONS = {
    "🏷️ 요정비닐 상품 현황": "pages/01_product_status.py",
    "🚚 밀크런 PPT 변환": "pages/02_milkrun_ppt.py",
    "📦 택배 송장 변환": "pages/03_invoice.py",
    "🏭 원가 시뮬레이터": "pages/04_cost_simulator.py",
    "📈 시장 지표 분석": "pages/05_market_index.py",
    "🏛️ 나라장터 입찰": "pages/06_narajangte.py"
}

# 라디오 버튼 생성
selection = st.sidebar.radio("메뉴를 선택하세요", list(MENU_OPTIONS.keys()))

# 5. 페이지 전환 로직 (선택 시 해당 파일 실행)
# st.switch_page는 스트림릿 1.30 버전 이상에서 권장하는 방식입니다.
if selection:
    st.switch_page(MENU_OPTIONS[selection])
