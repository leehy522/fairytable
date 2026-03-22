import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from auth import check_password

# 1. 페이지 설정 및 보안 (최상단)
st.set_page_config(page_title="요정비닐 상품 현황", page_icon="🏷️", layout="wide")

if not check_password():
    st.stop()

# 2. '웹에 게시' URL 폐기 및 안전한 편집용 URL 적용
# 주의: 이전에 원가 시뮬레이터에서 사용한 것과 동일한 스프레드시트 주소를 입력하십시오.
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU/edit?gid=0#gid=0"

@st.cache_data(ttl=60)
def _load_data() -> pd.DataFrame:
    """GCP 서비스 계정을 통해 구글 시트를 안전하게 읽어옵니다."""
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        # '상품DB' 워크시트를 명시하여 불러옵니다.
        df = conn.read(spreadsheet=SPREADSHEET_URL, worksheet="상품DB", ttl=0)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(subset=["상품명"])
    except Exception as e:
        st.error(f"구글 시트 보안 연결 오류: {e}")
        return pd.DataFrame()

# 3. 메인 UI 실행부 (render 래퍼 제거)
st.title("🏷️ 요정비닐 상품 실시간 현황")
st.caption("구글 스프레드시트와 안전하게 동기화 중 (자동 갱신 주기: 1분)")
st.divider()

df = _load_data()

if df.empty:
    st.error("데이터를 불러올 수 없습니다. 스프레드시트 URL과 권한 설정을 확인해주세요.")
    st.stop()

col1, _ = st.columns([1, 1])
with col1:
    st.metric("총 등록 상품", f"{len(df)} 종")

with st.expander("🔍 상품 검색 및 필터", expanded=True):
    search_query = st.text_input("", placeholder="검색할 상품명을 입력하세요")

display_df = (
    df[df["상품명"].str.contains(search_query, na=False)] if search_query else df
)

st.subheader(f"📦 상품 목록 ({len(display_df)}건)")
st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "상품명": st.column_config.TextColumn("📋 상품명"),
        "현재판매가": st.column_config.NumberColumn("💰 단가", format="₩%d"),
    },
)
st.success("✅ 최신 데이터가 반영되었습니다.")
