import subprocess
import sys
import os
import streamlit as st
from auth import check_password
import importlib.util

# 2. 페이지 기본 설정
st.set_page_config(page_title="요정비닐 통합 시스템", page_icon="🏭", layout="wide")

# 3. 보안 인증 체크
if not check_password():
    st.stop()

# 4. 하위 폴더(pages)에 있는 agent_logic 모듈 안전하게 불러오기
try:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    logic_path = os.path.join(base_dir, "pages", "agent_logic.py")
    
    spec = importlib.util.spec_from_file_location("agent_logic", logic_path)
    agent_logic = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(agent_logic)
    analyze_order_data = agent_logic.analyze_order_data
except Exception as e:
    # 만약 agent_logic 파일을 아직 안 만드셨을 경우를 대비한 대체 함수
    def analyze_order_data(api_key, prompt):
        return f"⚠️ pages/agent_logic.py 파일을 불러오지 못했습니다. 오류: {e}"

# 5. 강제 라우팅 메뉴 구성
MENU_MAP = {
    "🏠 홈": None,
    "🤖 AI 발주 에이전트": "ai_agent",  # 메인 화면에서 챗봇으로 구동
    "🏷️ 상품 실시간 현황": "product_status.py",
    "📦 택배 송장 변환": "invoice.py",
    "🏭 원가 시뮬레이터": "cost_simulator.py",
    "📈 시장 지표 분석": "market_index.py",
    "🚚 밀크런 PPT 변환": "milkrun_ppt.py",
    "🏛️ 나라장터 입찰": "narajangte.py",
}

# 6. 강제 사이드바 렌더링
st.sidebar.title("🚀 요정비닐 관리자")
st.sidebar.markdown(f"**접속자:** 관리자")
st.sidebar.divider()

selection = st.sidebar.radio("메뉴 이동", list(MENU_MAP.keys()))

st.sidebar.divider()
if st.sidebar.button("🔒 시스템 로그아웃"):
    st.session_state.password_correct = False
    st.rerun()

# 7. 동적 페이지 로딩 함수 (pages/ 폴더 내 파일 실행용)
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

# 8. 메인 화면 조건부 렌더링
if selection == "🏠 홈":
    st.title("🚀 요정비닐 통합 관리 대시보드")
    st.success("✅ 시스템 인증이 완료되었습니다. 안전한 업무 환경입니다.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("💡 **안내**\n좌측 사이드바에서 원하는 업무 메뉴를 선택하십시오. 각 메뉴는 독립된 모듈로 작동합니다.")
    with col2:
        st.warning("⚠️ **주의**\n업무 종료 시 반드시 로그아웃을 클릭하여 세션을 종료하십시오.")

elif selection == "🤖 AI 발주 에이전트":
    st.title("🤖 요정비닐 AI 발주 분석 및 에이전트 챗봇")
    st.caption("제미나이 3.8-flash 두뇌를 활용하여 발주서 상품번호 대조 및 재고 분석을 수행합니다.")
    
    # 대화 기록 세션 초기화
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {"role": "model", "content": "안녕하세요 대표님! 요정비닐 AI 에이전트입니다. 오늘 검토할 발주 데이터나 상품번호 대조 요청을 입력해 주세요."}
        ]

    # API 키 입력 및 초기화 버튼
    col_k1, col_k2 = st.columns([3, 1])
    with col_k1:
        gemini_api_key = st.text_input("Gemini API Key 입력 (AIza...)", type="password", help="발급받으신 제미나이 API 키를 입력해주세요.")
    with col_k2:
        if st.button("대화 내용 초기화"):
            st.session_state.messages = [{"role": "model", "content": "대화 기록이 초기화되었습니다. 새로운 지시를 내려주세요!"}]
            st.rerun()

    st.divider()

    # 채팅 UI 출력
    for msg in st.session_state.messages:
        with st.chat_message("ai" if msg["role"] == "model" else "user"):
            st.markdown(msg["content"])

    # 사용자 입력 처리
    if prompt := st.chat_input("에이전트에게 지시할 내용 (예: 최근 발주서 상품번호 누락 검증해줘)"):
        if not gemini_api_key:
            st.error("상단 입력란에 Gemini API Key를 먼저 입력해 주세요!")
        else:
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("ai"):
                with st.spinner("요정 에이전트가 발주 데이터를 분석 중입니다..."):
                    # pages/agent_logic.py에 있는 함수 호출
                    ai_response = analyze_order_data(gemini_api_key, prompt)
                    st.markdown(ai_response)
                    st.session_state.messages.append({"role": "model", "content": ai_response})

else:
    # 선택된 메뉴의 파일 실행 (기존 pages/ 폴더 연동)
    load_page(MENU_MAP[selection])
