"""
요정비닐 스마트 시스템 - 메인 진입점
메뉴별 모듈 구조:
  pages/
    01_product_status.py   → 🏷️ 요정비닐 상품 현황
    02_milkrun_ppt.py      → 🚚 밀크런 PPT 변환
    03_invoice.py          → 📦 택배 송장 변환
    04_cost_simulator.py   → 🏭 원가 시뮬레이터
    05_market_index.py     → 📈 시장 지표 분석
    06_narajangte.py       → 🏛️ 나라장터 입찰
"""

import streamlit as st
from auth import check_password

# 반드시 가장 먼저 호출
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 사이드바 완전 숨기기
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none; }
        [data-testid="collapsedControl"] { display: none; }
    </style>
""", unsafe_allow_html=True)

# 로그인 체크
if not check_password():
    st.stop()

# 메뉴 import (로그인 통과 후)
from pages import (
    product_status,
    milkrun_ppt,
    invoice,
    cost_simulator,
    market_index,
    narajangte,
)

# ── 사이드바 ──────────────────────────────────────────────
st.sidebar.title("🚀 요정비닐 관리자")
MENU_MAP = {
    "🏷️ 요정비닐 상품 현황" : product_status,
    "🚚 밀크런 PPT 변환"    : milkrun_ppt,
    "📦 택배 송장 변환"     : invoice,
    "🏭 원가 시뮬레이터"    : cost_simulator,
    "📈 시장 지표 분석"     : market_index,
    "🏛️ 나라장터 입찰"     : narajangte,
}

menu = st.sidebar.radio("메뉴를 선택하세요", list(MENU_MAP.keys()))

# ── 선택된 메뉴 렌더링 ────────────────────────────────────
MENU_MAP[menu].render()
