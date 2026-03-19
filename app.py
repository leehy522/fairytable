import streamlit as st
from auth import check_password
import importlib.util

# 1. 페이지 설정
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 2. 로그인 체크
if not check_password():
    st.stop()

# 3. 메뉴 설정 (모든 파일이 app.py와 같은 위치에 있을 때)
# 💡 경로에서 'pages/'나 'views/'를 싹 제거했습니다.
MENU_MAP = {
    "🏷️ 요정비닐 상품 현황": "product_status.py",
    "🚚 밀크런 PPT 변환": "milkrun_ppt.py",
    "📦 택배 송장 변환": "invoice.py",
    "🏭 원가 시뮬레이터": "cost_simulator.py",
    "📈 시장 지표 분석": "market_index.py",
    "🏛️ 나라장터 입찰": "narajangte.py"
}

st.sidebar.title("🚀 요정비닐 관리자")
selection = st.sidebar.radio("메뉴를 선택하세요", list(MENU_MAP.keys()))

# 4. 선택된 파일 실행 로직
if selection:
    file_name = MENU_MAP[selection]
    try:
        # 같은 폴더에 있는 파일을 모듈로 읽어서 실행합니다.
        spec = importlib.util.spec_from_file_location("module.name", file_name)
        page_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(page_module)
        
        # 만약 각 파일 안에 render() 함수를 만들어두셨다면 아래처럼 호출도 가능합니다.
        # page_module.render() 
        
    except Exception as e:
        st.error(f"❌ '{file_name}' 실행 중 오류가 발생했습니다.")
        st.exception(e)
