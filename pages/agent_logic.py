from google import genai
from github import Github
import streamlit as st
import time
import datetime

def get_chat_session(api_key):
    """
    요정비닐 시스템 고도화 및 코드 진단을 위한 시니어 아키텍트 세션 초기화
    """
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


def analyze_order_data(api_key, user_prompt):
    """
    시스템 점검, 코드 리뷰, 고도화 요청을 대화 맥락과 함께 처리하는 챗봇 함수
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


def generate_and_create_pr(api_key, github_token, task_description):
    """
    AI가 지시사항에 맞는 코드를 작성하고, 깃허브에 새 브랜치를 파서 커밋한 뒤 PR을 생성합니다.
    """
    try:
        # 1. 제미나이로 최적화 코드 생성
        client = genai.Client(api_key=api_key)
        prompt = (
            f"요정비닐 ERP 시스템 고도화 작업 지시사항:\n> {task_description}\n\n"
            "위 요구사항을 해결하기 위한 Streamlit 기반의 완성도 높은 파이썬 코드를 작성해주세요. "
            "마크다운 코드 블록(```python 등)을 제외하고, **순수 파이썬 코드만** 출력하세요."
        )
        response = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt
        )
        new_code = response.text.strip().replace("```python", "").replace("```", "").strip()

        # 2. 깃허브 연동 및 PR 상신
        g = Github(github_token)
        repo = g.get_repo("leehy522/fairytable")

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        new_branch = f"feature/ai-architect-{timestamp}"

        base_ref = repo.get_git_ref("heads/develop")
        repo.create_git_ref(ref=f"refs/heads/{new_branch}", sha=base_ref.object.sha)

        file_path = f"pages/optimized_task_{timestamp}.py"
        commit_msg = f"🤖 AI System Architect: 시스템 고도화 코드 자동 생성 ({file_path})"
        repo.create_file(file_path, commit_msg, new_code, branch=new_branch)

        pr_title = f"🤖 [AI Architect PR] 시스템 고도화 요청 - {timestamp}"
        pr_body = f"### 📝 작업 지시사항\n> {task_description}\n\nAI 시스템 아키텍트가 웹 대시보드 자율 코딩 기능을 통해 생성한 코드 개선 PR입니다."
        
        pr = repo.create_pull(title=pr_title, body=pr_body, head=new_branch, base="develop")
        
        return True, pr.html_url, new_code

    except Exception as e:
        return False, str(e), ""
