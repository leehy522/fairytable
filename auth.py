"""
auth.py — 로그인 인증 모듈
아이디/비밀번호를 세션으로 관리합니다.
"""

import streamlit as st
import os

# 환경변수 또는 Streamlit Secrets에서 읽기
_USER_ID = st.secrets.get("USER_ID", os.getenv("USER_ID", "lhy"))
_USER_PW = st.secrets.get("USER_PW", os.getenv("USER_PW", ""))


def check_password() -> bool:
    """
    로그인 여부를 확인합니다.
    - 이미 로그인 됐으면 True 반환
    - 아니면 로그인 폼을 그린 후 False 반환
    """
    if st.session_state.get("password_correct"):
        return True

    # 로그인 화면에서 사이드바 숨기기
    st.markdown("""
        <style>
            [data-testid="stSidebar"] { display: none ; }
            [data-testid="collapsedControl"] { display: none ; }
        </style>
    """, unsafe_allow_html=True)

    # ... 이하 기존 코드 동일

    st.title("🔐 요정비닐 시스템 접속")
    col_l, _ = st.columns([1, 2])

    with col_l:
        input_id = st.text_input("ID", placeholder="아이디 입력", key="login_id")
        input_pw = st.text_input(
            "Password", type="password", placeholder="비밀번호 입력", key="login_pw"
        )
        if st.button("로그인 실행"):
            if input_id == _USER_ID and input_pw == _USER_PW:
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("❌ 정보가 일치하지 않습니다.")

    return False
