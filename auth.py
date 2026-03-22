import streamlit as st

def check_password() -> bool:
    """
    로그인 여부를 확인합니다.
    - 이미 로그인 됐으면 True 반환
    - 아니면 로그인 폼을 그린 후 False 반환
    """
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    # 1. 로그인 성공 상태: 사이드바 원상 복구 CSS 주입
    if st.session_state["password_correct"]:
        st.markdown("""
            <style>
                [data-testid="stSidebar"] { display: unset !important; }
                [data-testid="collapsedControl"] { display: unset !important; }
            </style>
        """, unsafe_allow_html=True)
        return True

    # 2. 보안 키 검증 (Fail-Closed)
    try:
        _USER_ID = st.secrets["USER_ID"]
        _USER_PW = st.secrets["USER_PW"]
    except KeyError:
        st.error("⚠️ 서버에 보안 자격 증명(secrets.toml)이 설정되지 않아 시스템을 잠급니다.")
        return False

    # 3. 로그인 대기 상태: 사이드바 원천 차단 CSS 주입
    st.markdown("""
        <style>
            [data-testid="stSidebar"] { display: none !important; }
            [data-testid="collapsedControl"] { display: none !important; }
        </style>
    """, unsafe_allow_html=True)

    st.title("🔐 요정비닐 시스템 접속")
    col_l, _ = st.columns([1, 2])

    with col_l:
        # st.form 유지 (엔터키 로그인 지원)
        with st.form("login_form"):
            input_id = st.text_input("ID", placeholder="아이디 입력")
            input_pw = st.text_input("Password", type="password", placeholder="비밀번호 입력")
            
            submit = st.form_submit_button("로그인 실행")

            if submit:
                if input_id == _USER_ID and input_pw == _USER_PW:
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("❌ 정보가 일치하지 않습니다.")

    return False
