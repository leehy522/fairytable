from datetime import datetime, timedelta
import pandas as pd
import requests
import streamlit as st

# ── API 설정 ──────────────────────────────────────────────
# 💡 실제 공공데이터포털에서 발급받은 본인의 인증키인지 다시 확인하세요!
_AUTH_KEY = "9542280dba7856322b0e5c72c63c510c1fb83bc06c8d62eccab4f58324646cfd"

_BID_URLS = {
    "물품": "http://apis.data.go.kr/1230000/BidPublicInfoService05/getBidPblancListInfoThng03",
    "용역": "http://apis.data.go.kr/1230000/BidPublicInfoService05/getBidPblancListInfoServ03",
}
_RESULT_URLS = {
    "물품": "http://apis.data.go.kr/1230000/BidPublicInfoService05/getOpengResultListInfoThng03",
    "용역": "http://apis.data.go.kr/1230000/BidPublicInfoService05/getOpengResultListInfoServ03",
}

def _make_params(keyword: str, start_dt: str, end_dt: str, rows: int) -> dict:
    return {
        "serviceKey" : _AUTH_KEY,
        "bidNtceNm"  : keyword,
        "type"       : "json",
        "numOfRows"  : str(rows),
        "pageNo"     : "1",
        "inqryDiv"   : "1",
        "inqryBgnDt" : start_dt + "0000", # API 규격에 따라 시분초 추가 필요할 수 있음
        "inqryEndDt" : end_dt + "2359",
    }

def _fetch_items(url_map: dict, keyword: str, start_dt: str, end_dt: str, rows: int) -> list[dict]:
    results = []
    for category, url in url_map.items():
        try:
            res = requests.get(
                url,
                params=_make_params(keyword, start_dt, end_dt, rows),
                timeout=15,
            )
            if res.status_code == 200:
                data = res.json()
                items = data.get("response", {}).get("body", {}).get("items", [])
                # 공공데이터 API는 검색 결과가 1개일 때 리스트가 아닌 딕셔너리로 줄 때가 있습니다.
                if isinstance(items, dict): items = [items]
                
                for item in items or []:
                    item["구분"] = category
                    results.append(item)
        except Exception as e:
            st.warning(f"[{category}] API 연결 중: {e}")
    return results

def render() -> None:
    st.title("🏛️ 나라장터 통합 정보 센터")
    st.info("💡 실시간 입찰 공고와 낙찰 결과를 조회합니다.")

    # 검색 조건
    with st.expander("🔍 검색 조건 설정", expanded=True):
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            keyword = st.text_input("검색 키워드", value="비닐봉투")
        with col2:
            days_back = st.number_input("조회 기간(일)", min_value=1, max_value=30, value=7)
        with col3:
            rows = st.number_input("출력 개수", min_value=5, max_value=100, value=20)

    end_dt   = datetime.now().strftime("%Y%m%d")
    start_dt = (datetime.now() - timedelta(days=int(days_back))).strftime("%Y%m%d")

    tab1, tab2 = st.tabs(["📢 실시간 입찰 공고", "📊 낙찰(개찰) 결과"])
    
    with tab1:
        if st.button("🚀 공고 검색", key="bid_btn"):
            with st.spinner("데이터 수집 중..."):
                all_items = _fetch_items(_BID_URLS, keyword, start_dt, end_dt, rows)
            if all_items:
                df = pd.DataFrame(all_items)
                cols = {"구분":"구분", "bidNtceNm":"공고명", "bidNtceDt":"공고일시", "ntceInsttNm":"기관", "bidNtceUrl":"링크"}
                st.dataframe(df[[k for k in cols.keys() if k in df.columns]].rename(columns=cols), use_container_width=True, hide_index=True)
            else:
                st.warning("조회된 공고가 없습니다.")

    with tab2:
        if st.button("📊 결과 분석", key="res_btn"):
            with st.spinner("결과 분석 중..."):
                all_results = _fetch_items(_RESULT_URLS, keyword, start_dt, end_dt, rows)
            if all_results:
                df_res = pd.DataFrame(all_results)
