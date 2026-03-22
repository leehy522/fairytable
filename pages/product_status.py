import streamlit as st
import pandas as pd
from auth import check_password

# 페이지 기본 설정 (상단에 위치)
st.set_page_config(page_title="상품 실시간 현황", page_icon="🏷️", layout="wide")

def show_product_status():
    st.title("🏷️ 상품 실시간 현황")
    st.markdown("---")

    # 1. 시트 연결 시도
    try:
        # 라이브러리 임포트 에러를 방지하기 위해 내부 문자열 경로를 사용합니다.
        # 이 방식은 'No module named' 에러를 회피하는 데 효과적입니다.
        conn = st.connection("gsheets", type="streamlit_gsheets.gsheets_connection.GSheetsConnection")
        
        # 2. Secrets에서 URL 가져오기
        if "connections" not in st.secrets or "gsheets" not in st.secrets["connections"]:
            st.error("❌ Secrets 설정에 [connections.gsheets] 섹션이 없습니다.")
            return

        url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        
        # 3. 데이터 읽기 (캐시를 사용하여 성능 최적화)
        df = conn.read(spreadsheet=url, ttl="10m") 

        if df is not None and not df.empty:
            st.success(f"✅ 데이터를 성공적으로 불러왔습니다. (최근 업데이트: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')})")
            
            # 데이터 필터링 기능 (검색어)
            search_query = st.text_input("🔍 상품명 또는 규격 검색", "")
            if search_query:
                # 데이터프레임 내 모든 컬럼에서 검색어 포함 여부 확인
                df = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]

            # 데이터 테이블 출력
            st.dataframe(df, use_container_width=True, height=600)
            
            # 엑셀 다운로드 버튼
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 현재 화면 데이터 다운로드 (CSV)",
                data=csv,
                file_name="product_status.csv",
                mime="text/csv",
            )
        else:
            st.warning("⚠️ 불러온 데이터가 비어 있습니다. 구글 시트 내용을 확인하십시오.")

    except Exception as e:
        st.error(f"❌ 데이터를 불러오는 중 오류가 발생했습니다.")
        st.info(f"**오류 상세:** {e}")
        st.markdown("""
        **조치 방법:**
        1. `requirements.txt`에 `st-gsheets-connection`이 있는지 확인하십시오.
        2. Streamlit Cloud Settings의 **Secrets**에 시트 URL이 정확히 입력되었는지 확인하십시오.
        3. 구글 시트가 **'링크가 있는 모든 사용자에게 공개'** 상태인지 확인하십시오.
        """)

# 보안 인증 후 페이지 실행
if check_password():
    show_product_status()
