from google import genai
import streamlit as st

def get_chat_session(api_key):
    """
    제미나이 채팅 세션을 초기화하고 세션 상태에 저장하여 대화 맥락을 유지합니다.
    """
    # 스트림릿 session_state에 이미 채팅 세션이 있다면 재사용
    if "gemini_chat_session" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        try:
            client = genai.Client(api_key=api_key)
            
            # 요정비닐 ERP 에이전트 페르소나 설정 (시스템 지시어)
            system_instruction = (
                "당신은 요정비닐(Fairy Dream Co., Ltd.)의 스마트 팩토리 ERP AI 에이전트입니다. "
                "대표님과 실시간으로 대화하며 발주서 데이터 분석, 상품번호 대조, 재고 검증 및 업무 지시를 수행합니다. "
                "이전 대화의 맥락과 내용을 정확히 기억하고 이어서 답변해주세요."
            )
            
            # 최신 3.8-flash 모델을 활용한 채팅 세션 생성
            chat = client.chats.create(
                model='gemini-3.8-flash',
                config={
                    "system_instruction": system_instruction,
                    "temperature": 0.3, # 실무 정확도를 위해 다소 차분한 답변 유도
                }
            )
            
            st.session_state.gemini_chat_session = chat
            st.session_state.current_api_key = api_key
        except Exception as e:
            return None
            
    return st.session_state.gemini_chat_session


def analyze_order_data(api_key, user_prompt):
    """
    사용자의 메시지를 제미나이 채팅 세션에 전달하고 기억된 맥락을 바탕으로 답변을 받아옵니다.
    """
    try:
        chat_session = get_chat_session(api_key)
        if not chat_session:
            return "❌ 제미나이 API 연결 실패: API 키를 확인해주세요."
            
        # 챗봇처럼 이전 대화 맥락을 유지하며 메시지 전송
        response = chat_session.send_message(user_prompt)
        return response.text
        
    except Exception as e:
        return f"❌ AI 대화 중 오류 발생: {e}"
