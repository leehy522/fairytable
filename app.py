import streamlit as st
from auth import check_password

# 1. 페이지 기본 설정 (가장 먼저 실행되어야 함)
st.set_page_config(page_title="요정비닐 스마트 시스템", page_icon="🏭", layout="wide")

# 2. 보안: 로그인 체크 및 미인증 시 사이드바 원천 차단
if not check_password():
    st.markdown("""
        <style>
            [data-testid="collapsedControl"] { display: none; }
            [data-testid="stSidebar"] { display: none; }
        </style>
    """, unsafe_allow_html=True)
    st.stop() # 로그인이 안 되면 여기서 코드 실행을 완전히 멈춤

# 3. 로그인 성공 시 노출되는 메인 홈 화면
st.title("🚀 요정비닐 통합 대시보드")
st.markdown("---")
st.success("✅ 관리자 인증이 완료되었습니다.")
st.info("👈 좌측 사이드바 메뉴를 열어 원하는 업무를 선택하십시오.")

# 추가적인 대시보드 요약 지표나 공지사항을 넣기에 적합한 공간입니다.
