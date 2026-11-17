from google import genai
import streamlit as st
import time

def get_chat_session(api_key):
    if "gemini_client" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        try:
            client = genai.Client(api_key=api_key)
            st.session_state.gemini_client = client
            st.session_state.current_api_key = api_key
            
            system_instruction = (
                "당신은 요정비닐(Fairy Dream Co., Ltd.)의 스마트 팩토리 ERP AI 에이전트입니다. "
                "대표님과 실시간으로 대화하며 발주서 데이터 분석, 상품번호 대조, 재고 검증 및 업무 지시를 수행합니다. "
                "이전 대화의 맥락과 내용을 정확히 기억하고 이어서 답변해주세요."
            )
            
            chat = client.chats.create(
                model='gemini-3.8-flash',
                config={
                    "system_instruction": system_instruction,
                    "temperature": 0.3,
                }
            )
            st.session_state.gemini_chat_session = chat
        except Exception as e:
            st.session_state.init_error = str(e)
            return None
            
    return st.session_state.get("gemini_chat_session")


def analyze_order_data(api_key, user_prompt):
    if "init_error" in st.session_state and st.session_state.init_error:
        err = st.session_state.init_error
        st.session_state.init_error = None
        return f"❌ [초기화 실패 에러]: {err}"

    chat_session = get_chat_session(api_key)
    if not chat_session:
        return "❌ 제미나이 API 연결 실패: API 키를 다시 확인해주세요."
        
    # 지수 백오프(Exponential Backoff) 재시도 로직
    # 시도 횟수가 늘어날수록 대기 시간을 2초 -> 4초로 늘려 안정성 확보
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
                wait_time *= 2  # 대기 시간 배증
                continue
            else:
                return f"❌ [API 통신 중 에러 발생]: {error_str}"
