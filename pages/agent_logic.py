from google import genai
import streamlit as st
import time

def get_chat_session(api_key):
    """
    시스템 고도화 및 코드 진단을 위한 시니어 소프트웨어 아키텍트 페르소나 설정
    """
    if "gemini_client" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        try:
            client = genai.Client(api_key=api_key)
            st.session_state.gemini_client = client
            st.session_state.current_api_key = api_key
            
            # 고도화 에이전트 전용 시스템 인스트럭션
            system_instruction = (
                "당신은 요정비닐(Fairy Dream Co., Ltd.)의 스마트 팩토리 ERP 시스템을 진단하고 "
                "효율을 극한으로 고도화하는 수석 시니어 시스템 아키텍트이자 AI 에이전트입니다. "
                "Streamlit 대시보드 구조, 파이썬 코드 최적화, 업무 프로세스 자동화, 에러 핸들링, "
                "그리고 UI/UX 개선 방안에 대해 전문적이고 실효성 있는 코드와 솔루션을 제시합니다. "
                "대표님이 제공하는 코드나 시스템 상황을 면밀히 분석하고 효율적인 개선안을 도출해주세요."
            )
            
            chat = client.chats.create(
                model='gemini-3.8-flash',
                config={
                    "system_instruction": system_instruction,
                    "temperature": 0.2, # 코드 분석의 정확도를 높이기 위해 낮게 설정
                }
            )
            st.session_state.gemini_chat_session = chat
        except Exception as e:
            st.session_state.init_error = str(e)
            return None
            
    return st.session_state.get("gemini_chat_session")


def analyze_order_data(api_key, user_prompt):
    """
    시스템 점검 및 고도화 관련 질문/코드 리뷰 요청을 처리하는 함수
    """
    if "init_error" in st.session_state and st.session_state.init_error:
        err = st.session_state.init_error
        st.session_state.init_error = None
        return f"❌ [초기화 실패 에러]: {err}"

    chat_session = get_chat_session(api_key)
    if not chat_session:
        return "❌ 제미나이 API 연결 실패: API 키를 다시 확인해주세요."
        
    max_retries = 3
    wait_time = 2
    
    for attempt in range(max_retries):
        try:
            response = chat_session.send_message(user_prompt)
            return response.text
        except Exception as e:
            error_str = str(e)
            if ("503" in error_str or "UNAVAILABLE" in error_str) and attempt < max_retries - 1:
                time.sleep(wait_time)
                wait_time *= 2
                continue
            else:
                return f"❌ [API 통신 중 에러 발생]: {error_str}"
