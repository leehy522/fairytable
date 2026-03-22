import streamlit as st
from auth import check_password

st.set_page_config(page_title="요정비닐 시스템", layout="wide")

# 여기서 로그인 체크를 호출하면, 사이드바 숨김/노출 여부는 auth.py가 알아서 통제합니다.
if not check_password():
    st.stop()

st.title("🚀 요정비닐 통합 대시보드")
st.success("✅ 관리자 인증이 완료되었습니다.")
st.info("👈 좌측 사이드바 메뉴를 열어 원하는 업무를 선택하십시오.")
