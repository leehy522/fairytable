from google import genai
import streamlit as st

def get_chat_session(api_key):
    """
    제미나이 클라이언트와 채팅 세션을 session_state에 함께 저장하여 
    클라이언트가 닫히지 않고 대화 맥락이 유지되도록 합니다.
    """
    # 세션에 클라이언트가 없거나 API 키가 변경된 경우에만 새로 생성
    if "gemini_client" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        try:
            # 클라이언트 자체를 session_state에 보관하여 세션 만료 및 닫힘 방지
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
            return None
            
    return st.session_state.gemini_chat_session


def analyze_order_data(api_key, user_prompt):
    """
    유지되는 클라이언트 세션을 통해 메시지를 전송하고 답변을 받아옵니다.
    """
    try:
        chat_session = get_chat_session(api_key)
        if not chat_session:
            return "❌ 제미나이 API 연결 실패: API 키를 확인해주세요."
            
        response = chat_session.send_message(user_prompt)
        return response.text
        
    except Exception as e:
        return f"❌ AI 대화 중 오류 발생: {e}"
