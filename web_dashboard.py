import streamlit as st
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt

# 웹 페이지 설정
st.set_page_config(page_title="요정비닐 스마트 대시보드", layout="wide")
st.title("📊 요정비닐 경영 분석 시스템 (v1.0)")
st.sidebar.header("조회 설정")

# 데이터 가져오기 함수
def get_data():
    symbols = {"WTI 유가": "CL=F", "원/달러 환율": "KRW=X"}
    df = pd.DataFrame()
    for name, sym in symbols.items():
        data = yf.download(sym, period="1y", interval="1d")
        df[name] = data['Close']
    return df.ffill()

# 실행 버튼
if st.sidebar.button("데이터 업데이트"):
    with st.spinner('최신 정보를 가져오는 중...'):
        df = get_data()
        
        # 상단 지표 (Metric) 표시
        col1, col2 = st.columns(2)
        current_oil = df["WTI 유가"].iloc[-1]
        current_krw = df["원/달러 환율"].iloc[-1]
        
        col1.metric("현재 WTI 유가", f"${current_oil:.2f}")
        col2.metric("현재 원/달러 환율", f"₩{current_krw:.2f}")

        # 그래프 시각화
        fig, ax1 = plt.subplots(figsize=(10, 5))
        ax2 = ax1.twinx()
        
        ax1.plot(df.index, df["WTI 유가"], color='tab:blue', label='WTI')
        ax2.plot(df.index, df["원/달러 환율"], color='tab:red', label='환율', linestyle='--')
        
        st.pyplot(fig)

        # 데이터 표 출력
        st.subheader("📋 최근 시장 데이터")
        st.dataframe(df.tail(10))