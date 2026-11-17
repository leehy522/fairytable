from google import genai
from github import Github
import streamlit as st
import time

def get_chat_session(api_key):
    if "gemini_client" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        try:
            client = genai.Client(api_key=api_key)
            st.session_state.gemini_client = client
            st.session_state.current_api_key = api_key
            
            system_instruction = (
                "당신은 요정비닐(Fairy Dream Co., Ltd.)의 스마트 팩토리 ERP 시스템을 진단하고 "
                "효율을 극한으로 고도화하는 수석 시니어 시스템 아키텍트이자 AI 에이전트입니다. "
                "Streamlit 대시보드 구조, 파이썬 코드 최적화, 업무 프로세스 자동화, 에러 핸들링, "
                "그리고 UI/UX 개선 방안에 대해 전문적이고 실효성 있는 코드와 솔루션을 제시합니다."
            )
            
            chat = client.chats.create(
                model='gemini-3.8-flash',
                config={
                    "system_instruction": system_instruction,
                    "temperature": 0.2,
                }
            )
            st.session_state.gemini_chat_session = chat
        except Exception as e:
            st.session_state.init_error = str(e)
            return None
            
    return st.session_state.get("gemini_chat_session")


def analyze_order_data(api_key, user_prompt, github_token=None):
    """
    에이전트와 대화하며, 필요시 깃허브 저장소 구조를 함께 진단하는 함수
    """
    if "init_error" in st.session_state and st.session_state.init_error:
        err = st.session_state.init_error
        st.session_state.init_error = None
        return f"❌ [초기화 실패 에러]: {err}"

    chat_session = get_chat_session(api_key)
    if not chat_session:
        return "❌ 제미나이 API 연결 실패: API 키를 다시 확인해주세요."
        
    # 만약 사용자가 깃허브 토큰을 입력했고 구조 점검을 원한다면 저장소 정보를 프롬프트에 동적으로 주입
    context_data = ""
    if github_token and ("저장소" in user_prompt or "구조" in user_prompt or "코드" in user_prompt):
        try:
            g = Github(github_token)
            repo = g.get_repo("leehy522/fairytable")
            contents = repo.get_contents("", ref="develop")
            file_list = [c.path for c in contents]
            context_data = f"\n\n[현재 깃허브(develop) 저장소 파일 목록]: {file_list}"
        except Exception as e:
            context_data = f"\n\n[깃허브 연동 경고]: 토큰이 유효하지 않거나 접근 실패 ({e})"

    full_prompt = user_prompt + context_data

    max_retries = 3
    wait_time = 2
    
    for attempt in range(max_retries):
        try:
            response = chat_session.send_message(full_prompt)
            return response.text
        except Exception as e:
            error_str = str(e)
            if ("503" in error_str or "UNAVAILABLE" in error_str) and attempt < max_retries - 1:
                time.sleep(wait_time)
                wait_time *= 2
                continue
            else:
                return f"❌ [API 통신 중 에러 발생]: {error_str}"
