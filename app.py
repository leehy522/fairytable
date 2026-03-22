import streamlit as st
from auth import check_password

st.set_page_config(page_title="요정비닐 시스템", layout="wide")
import streamlit as st
from auth import check_password
import importlib.util
import os

# 1. 페이지 기본 설정
st.set_page_config(page_title="요정비닐 통합 시스템", page_icon="🏭", layout="wide")

# 2. 보안 인증 체크 (auth.py 활용)
if not check_password():
    st.stop()

# 3. 메뉴 구성 (파일명과 표시될 명칭 매핑)
# 파일이 pages 폴더 안에 있는지 반드시 확인하십시오.
MENU_MAP = {
    "🏛️ 나라장터 입찰": "narajangte.py",
    "🏷️ 상품 실시간 현황": "product_status.py",
    "📦 택배 송장 변환": "invoice.py",
    "🏭 원가 시뮬레이터": "cost_simulator.py",
    "📈 시장 지표 분석": "market_index.py",
    "🚚 밀크런 PPT 변환": "milkrun_ppt.py",
    "마진 계산기": "margin_calc.py",
}

# 4. 강제 사이드바 생성
st.sidebar.title("🚀 요정비닐 관리자")
st.sidebar.markdown("---")

# 기본 선택값을 '홈'으로 두고 싶다면 메뉴에 '🏠 홈'을 추가하는 것이 좋습니다.
menu_list = ["🏠 홈"] + list(MENU_MAP.keys())
selection = st.sidebar.radio("메뉴를 선택하세요", menu_list)

st.sidebar.markdown("---")
if st.sidebar.button("🔒 로그아웃"):
    st.session_state.password_correct = False
    st.rerun()

# 5. 동적 페이지 로딩 로직 (핵심)
def load_page(file_name):
    # app.py 위치를 기준으로 pages 폴더 내 파일 경로 추적
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(base_dir, "pages", file_name)
    
    try:
        spec = importlib.util.spec_from_file_location("page_module", file_path)
        page_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(page_module)
        # 각 페이지 파일에 render() 함수가 없어도 실행되도록 설계됨
    except FileNotFoundError:
        st.error(f"❌ 파일을 찾을 수 없습니다: {file_name}\n경로: {file_path}")
    except Exception as e:
        st.error(f"❌ 페이지 로딩 중 오류 발생: {e}")

# 6. 화면 출력부 (핵심 수정 구간)
if selection == "🏠 홈":
    # 홈 메뉴일 때만 대시보드 메인 제목을 보여줍니다.
    st.title("🚀 요정비닐 통합 대시보드")
    st.success("✅ 관리자 인증이 완료되었습니다.")
    st.info("👈 좌측 사이드바 메뉴를 열어 업무를 선택하십시오.")
    
    # 여기에 공장 가동 현황 요약이나 공지사항을 넣으면 완벽합니다.
    st.divider()
    st.write("현재 시스템 버전: v2.0 (모듈화 완료)")

else:
    # 홈이 아닌 다른 메뉴를 클릭했을 때는 '요정비닐 통합 대시보드' 제목 없이 
    # 해당 페이지의 내용만 깔끔하게 출력합니다.
    load_page(MENU_MAP[selection])
