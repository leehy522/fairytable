import subprocess
import sys
import os
import streamlit as st
from auth import check_password
import importlib.util

# 2. 페이지 기본 설정
st.set_page_config(page_title="요정비닐 통합 시스템", page_icon="🏭", layout="wide")

# 3. 보안 인증 체크 (사이드바 숨김/노출은 auth.py가 제어)
if not check_password():
    st.stop()

# 4. 강제 라우팅 메뉴 구성
# 파일이 pages/ 폴더 내에 정확히 존재해야 합니다.
MENU_MAP = {
    "🏠 홈": None,
    "🏷️ 상품 실시간 현황": "product_status.py",
    "📦 택배 송장 변환": "invoice.py",
    "🏭 원가 시뮬레이터": "cost_simulator.py",
    "📈 시장 지표 분석": "market_index.py",
    "🚚 밀크런 PPT 변환": "milkrun_ppt.py",
    "🏛️ 나라장터 입찰": "narajangte.py",
}

# 5. 강제 사이드바 렌더링
st.sidebar.title("🚀 요정비닐 관리자")
st.sidebar.markdown(f"**접속자:** 관리자")
st.sidebar.divider()

selection = st.sidebar.radio("메뉴 이동", list(MENU_MAP.keys()))

st.sidebar.divider()
if st.sidebar.button("🔒 시스템 로그아웃"):
    st.session_state.password_correct = False
    st.rerun()

# 6. 동적 페이지 로딩 함수
def load_page(file_name):
    if not file_name: return
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, "pages", file_name)
    
    try:
        spec = importlib.util.spec_from_file_location("page_module", file_path)
        page_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(page_module)
    except Exception as e:
        st.error(f"❌ 페이지를 불러오는 중 오류가 발생했습니다.\n\n파일: {file_name}\n오류내용: {e}")

# 7. 메인 화면 조건부 렌더링
if selection == "🏠 홈":
    st.title("🚀 요정비닐 통합 관리 대시보드")
    st.success("✅ 시스템 인증이 완료되었습니다. 안전한 업무 환경입니다.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("💡 **안내**\n좌측 사이드바에서 원하는 업무 메뉴를 선택하십시오. 각 메뉴는 독립된 모듈로 작동합니다.")
    with col2:
        st.warning("⚠️ **주의**\n업무 종료 시 반드시 로그아웃을 클릭하여 세션을 종료하십시오.")
else:
    # 선택된 메뉴의 파일 실행
    load_page(MENU_MAP[selection])
