from google import genai
from github import Github
import streamlit as st
import time
import datetime
import pandas as pd

# [설정] 구글 시트 CSV 웹 게시 링크
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU/export?format=csv"

def get_sheet_data():
    try:
        df = pd.read_csv(SHEET_CSV_URL, on_bad_lines='skip', engine='python')
        return df.to_string(index=False)
    except Exception as e:
        return f"시트 데이터 로드 실패: {e}"

def get_chat_session(api_key):
    if "gemini_client" not in st.session_state or st.session_state.get("current_api_key") != api_key:
        try:
            client = genai.Client(api_key=api_key)
            st.session_state.gemini_client = client
            st.session_state.current_api_key = api_key
            
            system_instruction = (
                "당신은 요정비닐(Fairy Dream Co., Ltd.)의 스마트 팩토리 ERP 시스템을 진단하고 "
                "효율을 극한으로 고도화하는 수석 시니어 시스템 아키텍트이자 AI 에이전트입니다. "
                "Streamlit 대시보드 구조, 파이썬 코드 최적화, 업무 프로세스 자동화뿐만 아니라, "
                "깃허브 저장소의 실제 소스 코드를 읽어 분석하고, 구글 시트의 실시간 상품 데이터를 결합하여 완벽한 솔루션을 제시합니다."
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
    if "init_error" in st.session_state and st.session_state.init_error:
        err = st.session_state.init_error
        st.session_state.init_error = None
        return f"❌ [초기화 실패 에러]: {err}"

    chat_session = get_chat_session(api_key)
    if not chat_session:
        return "❌ 제미나이 API 연결 실패: API 키를 다시 확인해주세요."
        
    # 1. 깃허브 저장소 코드 및 파일 구조 실시간 스캔 (시크릿 토큰 자동 활용)
    github_code_context = ""
    try:
        gh_token = st.secrets.get("GITHUB_TOKEN", "")
        if gh_token:
            g = Github(gh_token)
            repo = g.get_repo("leehy522/fairytable")
            
            # develop 브랜치 파일 목록 가져오기
            contents = repo.get_contents("", ref="develop")
            file_list = [c.path for c in contents]
            
            github_code_context += f"\n\n[현재 깃허브(develop) 저장소 파일 구조]: {file_list}"
            
            # 사용자가 특정 파일(예: app.py, invoice.py 등)에 대해 물어보면 해당 내용도 가져오기
            for file_name in ["app.py", "pages/invoice.py", "pages/cost_simulator.py"]:
                try:
                    f_content = repo.get_contents(file_name, ref="develop")
                    decoded_content = f_content.decoded_content.decode("utf-8")
                    # 내용이 너무 길면 일부 발췌 또는 전체 주입
                    github_code_context += f"\n\n--- [깃허브 파일 내용: {file_name}] ---\n{decoded_content[:3000]}..."
                except:
                    pass
    except Exception as e:
        github_code_context = f"\n\n[깃허브 연동 참고]: 코드 스캔 중 예외 발생 ({e})"

    # 2. 구글 시트 실시간 데이터 가져오기
    live_sheet_data = get_sheet_data()
    
    # 3. 모든 컨텍스트를 종합하여 프롬프트 구성
    full_prompt = (
        f"{user_prompt}\n"
        f"{github_code_context}\n\n"
        f"--- [구글 스프레드시트 실시간 상품 데이터] ---\n"
        f"{live_sheet_data}"
    )

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

def generate_and_create_pr(api_key, github_token, task_description):
    try:
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
