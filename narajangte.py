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
    if not st.button("🚀 검색 시작"):
        return

    with st.spinner(f"'{keyword}' 관련 공고를 찾는 중..."):
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
    raw_df = df[list(valid_cols.keys())].rename(columns=valid_cols)

    with st.expander("👀 나라장터 API 원본 데이터 보기 (필터링 전)", expanded=False):
        st.info(f"API가 가져온 총 {len(raw_df)}건의 날것 데이터입니다.")
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    # 스마트 필터링
    display_df = raw_df.copy()
    # 💡 [수정 후] 합집합 (OR 조건: 단어 중 하나라도 포함되면 합격)
    # 사용자가 "재활용 봉투"라고 치면 "재활용" 이거나 "봉투"가 하나라도 들어간 공고를 모두 찾습니다.
    if keyword.strip():
        or_pattern = '|'.join(keyword.split())
        display_df = display_df[display_df["공고명"].str.contains(or_pattern, na=False, regex=True)]

    st.success(f"✅ 필터링 완료: 총 {len(display_df)}건의 정확한 공고를 찾았습니다.")
    
    # 💡 [핵심] 물품과 용역을 분리합니다
    df_thng = display_df[display_df["구분"] == "물품"]
    df_serv = display_df[display_df["구분"] == "용역"]

    # 물품 표 렌더링
    st.subheader(f"📦 물품 공고 ({len(df_thng)}건)")
    st.dataframe(
        df_thng,
        use_container_width=True,
        hide_index=True,
        column_config={"상세링크": st.column_config.LinkColumn("상세링크", display_text="링크가기 🔗")}
    )

    # 용역 표 렌더링
    st.subheader(f"🛠️ 용역 공고 ({len(df_serv)}건)")
    st.dataframe(
        df_serv,
        use_container_width=True,
        hide_index=True,
        column_config={"상세링크": st.column_config.LinkColumn("상세링크", display_text="링크가기 🔗")}
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
    
    # 💡 [핵심 안전장치] 공고번호가 없거나 비어있는 경우를 완벽하게 걸러냅니다!
    if "bidNtceNo" in df_res.columns:
        # 공고번호가 정상적으로 존재할 때만 링크를 조립합니다.
        df_res["bidNtceUrl"] = df_res["bidNtceNo"].apply(
            lambda x: f"https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do?bidno={x}" if pd.notnull(x) else "https://www.g2b.go.kr"
        )
    else:
        # 공고번호 컬럼 자체가 없으면 강제로 빈칸을 만들고 메인 홈피로 연결합니다.
        df_res["bidNtceUrl"] = "https://www.g2b.go.kr"
        df_res["bidNtceNo"] = None

    res_cols = {
        "구분": "구분",
        "bidNtceNo": "공고번호",
        "bidNtceNm": "공고명",
        "opengDt": "개찰일시",
        "bidWinnerNm": "🏆 낙찰(1순위)업체",
        "totScor": "💯 종합점수",
        "tndrAmt": "💰 투찰금액(원)",
        "sucbidLwstRate": "📉 낙찰하한율(%)",
        "bidNtceUrl": "상세링크", 
    }
    
    # 컬럼이 빠져있을 경우 빈칸으로 강제 생성하여 에러 원천 차단
    for col_key in res_cols.keys():
        if col_key not in df_res.columns:
            df_res[col_key] = None

    raw_df_res = df_res[list(res_cols.keys())].rename(columns=res_cols)

    with st.expander("👀 나라장터 API 원본 데이터 보기 (필터링 전)", expanded=False):
        st.info(f"API가 가져온 총 {len(raw_df_res)}건의 날것 데이터입니다.")
        st.dataframe(raw_df_res, use_container_width=True, hide_index=True)

    # 스마트 필터링 (OR 조건)
    display_df = raw_df_res.copy()
    if keyword.strip():
        or_pattern = '|'.join(keyword.split())
        display_df = display_df[display_df["공고명"].str.contains(or_pattern, na=False, regex=True)]

    st.success(f"✅ 필터링 완료: 총 {len(display_df)}건의 정확한 낙찰 데이터를 분석했습니다.")
    
    df_thng = display_df[display_df["구분"] == "물품"]
    df_serv = display_df[display_df["구분"] == "용역"]

    col_config = {
        "💰 투찰금액(원)": st.column_config.NumberColumn("💰 투찰금액(원)", format="%d"),
        "💯 종합점수": st.column_config.NumberColumn("💯 종합점수", format="%.2f"),
        "상세링크": st.column_config.LinkColumn("상세링크", display_text="링크가기 🔗"),
    }

    st.subheader(f"📦 물품 낙찰 결과 ({len(df_thng)}건)")
    st.dataframe(df_thng, use_container_width=True, hide_index=True, column_config=col_config)

    st.subheader(f"🛠️ 용역 낙찰 결과 ({len(df_serv)}건)")
    st.dataframe(df_serv, use_container_width=True, hide_index=True, column_config=col_config)
    
# ── 메인 렌더링 ───────────────────────────────────────────

def render() -> None:
    st.title("🏛️ 나라장터 통합 정보 센터 (v1.7)")
    st.info("💡 하나의 인증키로 입찰 공고와 낙찰 결과를 모두 조회합니다.")

    # 검색 조건
    with st.expander("🔍 검색 조건 설정", expanded=True):
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            keyword = st.text_input("검색 키워드", value="비닐 봉투")
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
