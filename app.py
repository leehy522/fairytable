import streamlit as st
from auth import check_password
import importlib.util
import os

# 1. 페이지 설정
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 2. 로그인 전 사이드바 숨기기
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none; }
        [data-testid="collapsedControl"] { display: none; }
    </style>
""", unsafe_allow_html=True)

# 3. 로그인 체크
if not check_password():
    st.stop()

# 4. 로그인 후 사이드바 표시
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: block; }
        [data-testid="collapsedControl"] { display: block; }
    </style>
""", unsafe_allow_html=True)

# 5. 메뉴 설정
MENU_MAP = {
    "🏷️ 요정비닐 상품 현황" : "product_status.py",
    "🚚 밀크런 PPT 변환"    : "milkrun_ppt.py",
    "📦 택배 송장 변환"     : "invoice.py",
    "🏭 원가 시뮬레이터"    : "cost_simulator.py",
    "📈 시장 지표 분석"     : "market_index.py",
    "🏛️ 나라장터 입찰"     : "narajangte.py"
}

st.sidebar.title("🚀 요정비닐 관리자")
selection = st.sidebar.radio("메뉴를 선택하세요", list(MENU_MAP.keys()))

# 6. 현재 파일 기준 절대경로로 실행 ← 핵심 수정
if selection:
    # app.py가 있는 폴더를 기준으로 경로 설정
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, MENU_MAP[selection])

    try:
        spec = importlib.util.spec_from_file_location("page_module", file_path)
        page_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(page_module)
        if hasattr(page_module, "render"):
            page_module.render()
    except Exception as e:
        st.error(f"❌ '{MENU_MAP[selection]}' 실행 중 오류가 발생했습니다.")
        st.exception(e)
