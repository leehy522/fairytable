import os
import streamlit as st
from auth import check_password
import importlib.util
from pages.agent_logic import generate_and_create_pr, analyze_order_data

# 2. 페이지 기본 설정
st.set_page_config(page_title="요정비닐 통합 시스템", page_icon="🏭", layout="wide")

# 3. 보안 인증 체크
if not check_password():
    st.stop()

# 4. 강제 라우팅 메뉴 구성
MENU_MAP = {
    "🏠 홈": None,
    "🤖 AI 시스템 아키텍트": "ai_architect",
    "🏷️ 상품 실시간 현황": "product_status.py",
    "📦 택배 송장 변환": "invoice.py",
    "🏭 원가 시뮬레이터": "cost_simulator.py",
    "📈 시장 지표 분석": "market_index.py",
    "🚚 밀크런 PPT 변환": "milkrun_ppt.py",
    "🏛️ 나라장터 입찰": "narajangte.py",
}

# 5. 강제 사이드바 렌더링 (시크릿 자동 연동 + 수동 입력 지원)
st.sidebar.title("🚀 요정비닐 관리자")
st.sidebar.markdown(f"**접속자:** 관리자")
st.sidebar.divider()

st.sidebar.subheader("🔑 API 및 깃허브 인증")

# secrets.toml에서 키를 먼저 탐색하고, 없으면 빈 문자열 반환
default_gemini_key = st.secrets.get("GEMINI_API_KEY", "")
default_github_token = st.secrets.get("GITHUB_TOKEN", "")

gemini_api_key = st.sidebar.text_input(
    "Gemini API Key (AIza...)", 
    type="password", 
    value=default_gemini_key,
    help="secrets.toml에 설정된 경우 자동으로 로드됩니다."
)
github_token = st.sidebar.text_input(
    "GitHub Token (ghp_...)", 
    type="password", 
    value=default_github_token,
    help="secrets.toml에 설정된 경우 자동으로 로드됩니다."
)

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

elif selection == "🤖 AI 시스템 아키텍트":
    st.title("🤖 요정비닐 AI 시스템 아키텍트 & 자율 코딩 센터")
    st.caption("대화형 코드 진단부터 깃허브 자율 PR 상신까지 웹에서 원클릭으로 제어합니다.")
    
    sub_tab1, sub_tab2 = st.tabs(["💬 1:1 아키텍트 챗봇", "🚀 자율 코딩 & 깃허브 PR 상신"])
    
    # [탭 1] 실시간 챗봇
    with sub_tab1:
        st.subheader("💬 AI 시스템 아키텍트와 대화하기")
        
        if "messages" not in st.session_state:
            st.session_state.messages = [
                {"role": "model", "content": "안녕하세요 대표님! 요정비닐 시스템 아키텍트 에이전트입니다. 코드 개선안이나 시스템 구조 진단에 대해 말씀해 주세요."}
            ]

        for msg in st.session_state.messages:
            with st.chat_message("ai" if msg["role"] == "model" else "user"):
                st.markdown(msg["content"])

        if prompt := st.chat_input("아키텍트에게 문의할 내용을 입력하세요..."):
            if not gemini_api_key:
                st.error("사이드바에 Gemini API Key를 입력하거나 secrets.toml을 설정해 주세요!")
            else:
                st.session_state.messages.append({"role": "user", "content": prompt})
                with st.chat_message("user"):
                    st.markdown(prompt)

                with st.chat_message("ai"):
                    with st.spinner("아키텍트가 분석 중입니다..."):
                        ai_response = analyze_order_data(gemini_api_key, prompt)
                        st.markdown(ai_response)
                        st.session_state.messages.append({"role": "model", "content": ai_response})

    # [탭 2] 자율 코딩 및 깃허브 PR 상신
    with sub_tab2:
        st.subheader("🚀 AI 자율 코딩 및 깃허브 결재(PR) 자동 상신")
        st.caption("작업 지시를 내리면 에이전트가 코드를 직접 짜서 깃허브 저장소에 PR을 올립니다.")
        
        task_input = st.text_area(
            "실무 코드 고도화 작업 지시서",
            value="원가 시뮬레이터 모듈에 예외 처리 로직과 UI 디자인 개선 코드를 추가해줘."
        )
        
        if st.button("🚀 AI 자율 코딩 및 PR 생성 실행", type="primary"):
            if not gemini_api_key or not github_token:
                st.error("사이드바에 **Gemini API Key**와 **GitHub Token**이 모두 입력되어 있는지 확인해 주세요!")
            else:
                with st.spinner("AI 아키텍트가 코드를 작성하고 깃허브에 PR을 올리는 중입니다..."):
                    success, result_msg, generated_code = generate_and_create_pr(
                        gemini_api_key, github_token, task_input
                    )
                    
                    if success:
                        st.success("🎉 AI 아키텍트가 성공적으로 코드를 작성하고 결재 서류(PR)를 올렸습니다!")
                        st.markdown(f"👉 **생성된 깃허브 PR 확인하기:** [바로 가기 링크]({result_msg})")
                        
                        with st.expander("📝 에이전트가 작성한 실무 코드 보기"):
                            st.code(generated_code, language="python")
                    else:
                        st.error(f"❌ 깃허브 PR 생성 실패: {result_msg}")

else:
    load_page(MENU_MAP[selection])
