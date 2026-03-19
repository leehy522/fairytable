"""
pages/narajangte.py — 🏛️ 나라장터 입찰
공공데이터포털 API를 통해 입찰 공고 및 낙찰 결과를 조회합니다.
"""

from datetime import datetime, timedelta

import pandas as pd
import requests
import streamlit as st


# ── API 설정 ──────────────────────────────────────────────
_AUTH_KEY = "9542280dba7856322b0e5c72c63c510c1fb83bc06c8d62eccab4f58324646cfd"

# 1. 입찰 공고 최신 URL (숫자 03 제거 완료)
_BID_URLS = {
    "물품": "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThng",
    "용역": "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServc",
}

# 2. 낙찰 결과 최신 URL (숫자 03 제거 및 낙찰 전용 오퍼레이션으로 변경 완료)
_RESULT_URLS = {
    "물품": "https://apis.data.go.kr/1230000/as/ScsbidInfoService/getScsbidListSttusThng",
    "용역": "https://apis.data.go.kr/1230000/as/ScsbidInfoService/getScsbidListSttusServc",
}

def _make_params(keyword: str, start_dt: str, end_dt: str, rows: int) -> dict:
    return {
        "serviceKey" : _AUTH_KEY,
        "bidNtceNm"  : keyword,
        "type"       : "json",
        "numOfRows"  : str(rows),
        "pageNo"     : "1",
        "inqryDiv"   : "1", # 1: 날짜 기준 조회
        "inqryBgnDt" : start_dt + "0000", # API 문서 규격에 맞춘 12자리 (YYYYMMDDHHMM)
        "inqryEndDt" : end_dt + "2359",   # API 문서 규격에 맞춘 12자리 (YYYYMMDDHHMM)
    }

def _fetch_items(url_map: dict, keyword: str, start_dt: str,
                 end_dt: str, rows: int) -> list[dict]:
    results = []
    for category, url in url_map.items():
        try:
            # ✅ params 대신 URL 직접 조립 (이중 인코딩 방지)
            full_url = (
                f"{url}"
                f"?serviceKey={_AUTH_KEY}"
                f"&bidNtceNm={keyword}"
                f"&type=json"
                f"&numOfRows={rows}"
                f"&pageNo=1"
                f"&inqryDiv=1"
                f"&inqryBgnDt={start_dt}"
                f"&inqryEndDt={end_dt}"
            )
            res = requests.get(full_url, timeout=15)

            if res.status_code == 200:
                raw = res.json().get("response", {}).get("body", {}).get("items", [])
                if isinstance(raw, dict):
                    items = [raw]
                elif isinstance(raw, list):
                    items = raw
                else:
                    items = []
                for item in items:
                    item["구분"] = category
                results.extend(items)
            else:
                st.warning(f"[{category}] HTTP 오류: {res.status_code}")
        except Exception as e:
            st.warning(f"[{category}] API 오류: {e}")
    return results


# ── 탭 1: 입찰 공고 ───────────────────────────────────────

def _tab_bid_notice(keyword: str, start_dt: str, end_dt: str, rows: int) -> None:
    if not st.button("🚀 물품/용역 통합 검색"):
        return

    with st.spinner(f"'{keyword}' 관련 모든 공고를 찾는 중..."):
        all_items = _fetch_items(_BID_URLS, keyword, start_dt, end_dt, rows)

    if not all_items:
        st.warning(f"🔍 '{keyword}' 관련 공고가 물품/용역 모두에 없습니다.")
        return

    df = pd.DataFrame(all_items)
    cols = {
        "구분": "구분",
        "bidNtceNm": "공고명",
        "bidNtceDt": "공고일시",
        "ntceInsttNm": "공고기관",
        "bidNtceUrl": "상세링크",
    }
    valid_cols = {k: v for k, v in cols.items() if k in df.columns}
    
    # 원본 데이터프레임 (필터링 전)
    raw_df = df[list(valid_cols.keys())].rename(columns=valid_cols)

    # 💡 [추가됨] 원문 데이터를 접었다 펼칠 수 있는 공간
    with st.expander("👀 나라장터 API 원본 데이터 보기 (필터링 전)", expanded=False):
        st.info(f"API가 검색어와 조금이라도 연관 지어 가져온 총 {len(raw_df)}건의 날것 데이터입니다.")
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    # 띄어쓰기 기준 스마트 필터링 적용
    display_df = raw_df.copy()
    for kw in keyword.split():
        display_df = display_df[display_df["공고명"].str.contains(kw, na=False)]

    st.success(f"✅ 필터링 완료: 총 {len(display_df)}건의 정확한 공고를 찾았습니다.")
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "상세링크": st.column_config.LinkColumn(
                "상세링크",
                help="클릭하면 나라장터 공고 상세 페이지로 이동합니다",
                display_text="링크가기 🔗"
            )
        }
    )


# ── 탭 2: 낙찰 결과 ───────────────────────────────────────

def _tab_award_result(keyword: str, start_dt: str, end_dt: str, rows: int) -> None:
    if not st.button("📊 최근 낙찰 데이터 분석"):
        return

    with st.spinner(f"'{keyword}' 관련 낙찰 결과를 분석 중..."):
        all_results = _fetch_items(_RESULT_URLS, keyword, start_dt, end_dt, rows)

    if not all_results:
        st.warning(f"🔍 '{keyword}' 관련 최근 낙찰 결과가 없습니다.")
        return

    df_res = pd.DataFrame(all_results)
    res_cols = {
        "구분": "구분",
        "bidNtceNm": "공고명",
        "opengDt": "개찰일시",
        "bidWinnerNm": "🏆 낙찰(1순위)업체",
        "totScor": "💯 종합점수",
        "tndrAmt": "💰 투찰금액(원)",
        "sucbidLwstRate": "📉 낙찰하한율(%)",
    }
    
    for col_key in res_cols.keys():
        if col_key not in df_res.columns:
            df_res[col_key] = None

    # 원본 데이터프레임 (필터링 전)
    raw_df_res = df_res[list(res_cols.keys())].rename(columns=res_cols)

    # 💡 [추가됨] 원문 데이터를 접었다 펼칠 수 있는 공간
    with st

# ── 메인 렌더링 ───────────────────────────────────────────

def render() -> None:
    st.title("🏛️ 나라장터 통합 정보 센터 (v1.7)")
    st.info("💡 하나의 인증키로 입찰 공고와 낙찰 결과를 모두 조회합니다.")

    # 검색 조건
    with st.expander("🔍 검색 조건 설정", expanded=True):
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            keyword = st.text_input("검색 키워드", value="비닐봉투")
        with col2:
            days_back = st.number_input("조회 기간(일)", min_value=1, max_value=365, value=7)
        with col3:
            rows = st.number_input("출력 개수", min_value=5, max_value=1000, value=100)

    end_dt   = datetime.now().strftime("%Y%m%d") + "2359"
    start_dt = (datetime.now() - timedelta(days=int(days_back))).strftime("%Y%m%d") + "0000"

    tab1, tab2 = st.tabs(["📢 실시간 입찰 공고", "📊 낙찰(개찰) 결과"])
    with tab1:
        _tab_bid_notice(keyword, start_dt, end_dt, int(rows))
    with tab2:
        _tab_award_result(keyword, start_dt, end_dt, int(rows))
