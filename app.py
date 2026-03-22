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
selection = st.sidebar.radio("메뉴를 선택하세요", list(MENU_MAP.keys()))

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

# 선택된 메뉴 실행
if selection:
    load_page(MENU_MAP[selection])
# 여기서 로그인 체크를 호출하면, 사이드바 숨김/노출 여부는 auth.py가 알아서 통제합니다.
if not check_password():
    st.stop()

st.title("🚀 요정비닐 통합 대시보드")
st.success("✅ 관리자 인증이 완료되었습니다.")
st.info("👈 좌측 사이드바 메뉴를 열어 원하는 업무를 선택하십시오.")
