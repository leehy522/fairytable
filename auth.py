import streamlit as st

def check_password() -> bool:
    """
    로그인 여부를 확인합니다.
    - 이미 로그인 됐으면 True 반환
    - 아니면 로그인 폼을 그린 후 False 반환
    """
    # 1. 세션 상태 명시적 초기화
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    # 2. 강력한 보안 원칙 적용 (Fail-closed)
    # get() 대신 직접 키를 호출하여, 파일이 없으면 에러를 발생시키고 접근을 원천 차단합니다.
    try:
        _USER_ID = st.secrets["USER_ID"]
        _USER_PW = st.secrets["USER_PW"]
    except KeyError:
        st.error("⚠️ 서버에 보안 자격 증명(secrets.toml)이 설정되지 않아 시스템을 잠급니다.")
        return False

    # 3. 로그인 화면에서 사이드바 숨기기 (!important 추가로 강제성 부여)
    st.markdown("""
        <style>
            [data-testid="stSidebar"] { display: none !important; }
            [data-testid="collapsedControl"] { display: none !important; }
        </style>
    """, unsafe_allow_html=True)

    st.title("🔐 요정비닐 시스템 접속")
    col_l, _ = st.columns([1, 2])

    with col_l:
        # 4. st.form을 사용하여 'Enter' 키 입력 지원 및 불필요한 새로고침 방지
        with st.form("login_form"):
            input_id = st.text_input("ID", placeholder="아이디 입력")
            input_pw = st.text_input("Password", type="password", placeholder="비밀번호 입력")
            
            # form_submit_button을 사용해야 엔터키가 정상 작동합니다.
            submit = st.form_submit_button("로그인 실행")

            if submit:
                if input_id == _USER_ID and input_pw == _USER_PW:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("❌ 정보가 일치하지 않습니다.")

    return False
