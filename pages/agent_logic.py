from google import genai
import os

def analyze_order_data(api_key, user_prompt, file_data=None):
    """
    제미나이 3.8-flash 모델을 이용해 발주서 데이터 및 상품번호를 분석하는 함수
    """
    try:
        client = genai.Client(api_key=api_key)
        
        system_instruction = (
            "당신은 요정비닐(Fairy Dream Co., Ltd.)의 스마트 팩토리 ERP AI 에이전트입니다. "
            "제공된 발주 데이터와 상품번호를 정확하게 비교·대조하고, "
            "누락된 항목이나 수량 불일치 내역을 전문적이고 명확하게 리포트해주세요."
        )
        
        full_prompt = f"{system_instruction}\n\n사용자 지시사항:\n{user_prompt}"
        if file_data:
            full_prompt += f"\n\n참고 데이터:\n{file_data}"

        response = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=full_prompt
        )
        return response.text
    except Exception as e:
        return f"❌ AI 분석 중 오류 발생: {e}"
