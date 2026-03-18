"""
pages/market_index.py — 📈 시장 지표 분석
yfinance를 통해 WTI 유가 / 원달러 환율 2년치 데이터를 시각화합니다.
"""

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
import yfinance as yf


# ── 심볼 설정 ─────────────────────────────────────────────
_SYMBOLS = {
    "WTI 유가"   : "CL=F",
    "원/달러 환율": "KRW=X",
}


def _fetch_data() -> pd.DataFrame:
    """yfinance에서 2년치 종가 데이터를 가져옵니다."""
    df = pd.DataFrame()
    for name, sym in _SYMBOLS.items():
        data = yf.download(sym, period="2y", interval="1d")
        df[name] = data["Close"]
    return df.ffill()


def _draw_chart(df: pd.DataFrame) -> None:
    """이중 축 꺾은선 차트를 그립니다."""
    fig, ax1 = plt.subplots(figsize=(10, 5))
    ax2 = ax1.twinx()

    ax1.plot(df.index, df["WTI 유가"], color="tab:blue", label="WTI", linewidth=2)
    ax2.plot(
        df.index,
        df["원/달러 환율"],
        color="tab:red",
        label="환율",
        linestyle="--",
        linewidth=2,
    )

    ax1.set_ylabel("WTI Price (USD)", color="tab:blue")
    ax2.set_ylabel("Exchange Rate (KRW)", color="tab:red")
    plt.title("WTI Oil vs USD/KRW Exchange Rate")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left")

    st.pyplot(fig)


def render() -> None:
    st.title("📈 실시간 유가 및 환율 모니터링")
    st.write("WTI 유가와 원/달러 환율의 2년치 흐름을 실시간으로 가져옵니다.")

    if not st.button("📊 최신 데이터 불러오기"):
        return

    with st.spinner("데이터를 가져오는 중..."):
        try:
            df = _fetch_data()

            # 주요 지표
            c1, c2 = st.columns(2)
            c1.metric("현재 WTI 유가",  f"${df['WTI 유가'].iloc[-1]:.2f}")
            c2.metric("현재 환율",       f"₩{df['원/달러 환율'].iloc[-1]:.2f}")

            # 차트
            _draw_chart(df)

            # 상세 데이터 테이블
            st.subheader("📋 최근 데이터 상세")
            st.dataframe(
                df.tail(10).sort_index(ascending=False), use_container_width=True
            )

        except Exception as e:
            st.error(f"데이터 연동 실패: {e}")
