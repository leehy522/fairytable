from datetime import datetime, timedelta
import pandas as pd
import requests
import streamlit as st
from auth import check_password

# 1. 페이지 설정 및 보안 (최상단)
st.set_page_config(page_title="나라장터 입찰분석", page_icon="🏛️", layout="wide")

if not check_password():
    st.stop()

# ── API 설정 ──────────────────────────────────────────────
_AUTH_KEY = st.secrets["NARA_API_KEY"]

_BID_URLS = {
    "물품": "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThng",
    "용역": "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServc",
}

_RESULT_URLS = {
    "물품": "https://apis.data.go.kr/1230000/as/ScsbidInfoService/getScsbidListSttusThng",
    "용역": "https://apis.data.go.kr/1230000/as/ScsbidInfoService/getScsbidListSttusServc",
}

def _fetch_items(url_map: dict, keyword: str, start_dt: str,
                 end_dt: str, rows: int) -> list[dict]:
    results = []
    for category, url in url_map.items():
        try:
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
    if not st.button("🚀 검색 시작", key="btn_bid"):
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

    display_df = raw_df.copy()
    if keyword.strip():
        or_pattern = '|'.join(keyword.split())
        display_df = display_df[display_df["공고명"].str.contains(or_pattern, na=False, regex=True)]

    st.success(f"✅ 필터링 완료: 총 {len(display_df)}건의 정확한 공고를 찾았습니다.")
    
    df_thng = display_df[display_df["구분"] == "물품"]
    df_serv = display_df[display_df["구분"] == "용역"]

    st.subheader(f"📦 물품 공고 ({len(df_thng)}건)")
    st.dataframe(df_thng, use_container_width=True, hide_index=True, column_config={"상세링크": st.column_config.LinkColumn("상세링크", display_text="링크가기 🔗")})

    st.subheader(f"🛠️ 용역 공고 ({len(df_serv)}건)")
    st.dataframe(df_serv, use_container_width=True, hide_index=True, column_config={"상세링크": st.column_config.LinkColumn("상세링크", display_text="링크가기 🔗")})

# ── 탭 2: 낙찰 결과 ───────────────────────────────────────
def _tab_award_result(keyword: str, start_dt: str, end_dt: str, rows: int) -> None:
    if not st.button("📊 최근 낙찰 데이터 분석", key="btn_award"):
        return

    with st.spinner(f"'{keyword}' 관련 낙찰 결과를 분석 중..."):
        all_results = _fetch_items(_RESULT_URLS, keyword, start_dt, end_dt, rows)

    if not all_results:
        st.warning(f"🔍 '{keyword}' 관련 최근 낙찰 결과가 없습니다.")
        return

    df_res = pd.DataFrame(all_results)
    
    # 누락되었던 필수 데이터 컬럼 복구
    res_cols = {
        "구분": "구분",
        "bidNtceNo": "공고번호",
        "bidNtceNm": "공고명",
        "bidWinnerNm": "🏆 낙찰업체",
        "totScor": "💯 종합점수",
        "tndrAmt": "💰 투찰금액(원)",
        "sucbidLwstRate": "📉 낙찰하한율(%)",
        "bidNtceUrl": "상세링크", 
    }
    
    for col_key in res_cols.keys():
        if col_key not in df_res.columns:
            df_res[col_key] = None

    df_res["bidNtceUrl"] = "https://www.g2b.go.kr"
    raw_df_res = df_res[list(res_cols.keys())].rename(columns=res_cols)

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

# ── 메인 UI 실행부 (render 래퍼 제거) ────────────────────────
st.title("🏛️ 나라장터 통합 정보 센터 (v1.7)")
st.info("💡 하나의 인증키로 입찰 공고와 낙찰 결과를 모두 조회합니다.")

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
