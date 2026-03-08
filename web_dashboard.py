import streamlit as st
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import copy
import io
from pptx import Presentation
from pptx.util import Pt
from datetime import datetime

# --- [1. 페이지 기본 설정] ---
# set_page_config는 반드시 코드의 가장 상단(import 직후)에 한 번만 나와야 합니다.
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# --- [2. 공통 로직 및 함수 정의] ---
# (밀크런 관련 함수들: get_pallet_capacity, duplicate_slide, set_bold_text, fill_slide_data는 상단에 정의)
# (이미 정의된 함수들은 가독성을 위해 생략하며, 실제 코드에는 그대로 유지하시면 됩니다.)

@st.cache_data(ttl=60)
def load_google_sheet_data():
    CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTVvCbm9KEoUrqvlXSyIyLHmstIGZuiuTMLYDBnmgnxInrfoMelDXFSWogUdHUfNALb7uC_nBAIyzif/pub?output=csv"
    try:
        df = pd.read_csv(CSV_URL)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(subset=['상품명'])
    except Exception as e:
        st.error(f"구글 시트 연결 오류: {e}")
        return pd.DataFrame()

# --- [3. 사이드바 메뉴 구성] ---
st.sidebar.title("🚀 요정비닐 관리자")
menu = st.sidebar.radio("메뉴를 선택하세요", 
    ["🏷️ 요정비닐 상품 현황", "🚚 밀크런 PPT 변환", "📦 택배 송장 변환", "🏭 원가 시뮬레이터", "📈 시장 지표 분석"])

# --- [4. 메뉴별 화면 로직] ---

# --- 메뉴 1: 요정비닐 상품 현황 ---
if menu == "🏷️ 요정비닐 상품 현황":
    st.title("🏷️ 요정비닐 상품 실시간 현황")
    st.caption("구글 스프레드시트와 동기화 중 (자동 갱신 주기: 1분)")
    st.divider()

    df = load_google_sheet_data()

    if not df.empty:
        col1, col2 = st.columns([1, 1])
        with col1:
            st.metric("총 등록 상품", f"{len(df)} 종")
        
        with st.expander("🔍 상품 검색 및 필터", expanded=True):
            search_query = st.text_input("", placeholder="검색할 상품명을 입력하세요")
        
        display_df = df[df['상품명'].str.contains(search_query, na=False)] if search_query else df

        st.subheader(f"📦 상품 목록 ({len(display_df)}건)")
        st.dataframe(
            display_df, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "상품명": st.column_config.TextColumn("📋 상품명"),
                "단가": st.column_config.NumberColumn("💰 단가", format="₩%d"),
                "재고": st.column_config.NumberColumn("🔢 현재고", format="%d")
            }
        )
        st.success("✅ 최신 데이터가 반영되었습니다.")
    else:
        st.error("데이터를 불러올 수 없습니다. 구글 시트의 '웹에 게시' 설정을 확인해주세요.")

# --- 메뉴 2: 밀크런 PPT 변환 ---
elif menu == "🚚 밀크런 PPT 변환":
    st.title("🚚 밀크런 자동 변환 시스템")
    tpl_file = st.file_uploader("1. 밀크런_양식.pptx 업로드", type=['pptx'])
    pdf_files = st.file_uploader("2. 발주서 PDF 업로드", type=['pdf'], accept_multiple_files=True)
    
    if tpl_file and pdf_files:
        if st.button("🚀 PPT 생성 시작"):
            # (기존의 PPT 생성 로직 실행)
            st.success("변환 로직이 실행됩니다.")

# --- 메뉴 3: 택배 송장 변환 ---
elif menu == "📦 택배 송장 변환":
    st.title("📦 택배 송장 자동 변환기 (A-type)")
    col1, col2 = st.columns(2)
    with col1:
        input_file = st.file_uploader("1. 원본 주문 엑셀 선택", type=['xlsx', 'xls'])
    with col2:
        template_file = st.file_uploader("2. 템플릿 엑셀(A-type 양식) 선택", type=['xlsx', 'xls'])

    if input_file and template_file:
        if st.button("🚀 변환 실행"):
            # (기존의 엑셀 변환 및 다운로드 버튼 로직 실행)
            st.info("변환을 시작합니다.")

# --- 메뉴 4: 원가 시뮬레이터 ---
elif menu == "🏭 원가 시뮬레이터":
    st.title("🏭 원가 및 규격 시뮬레이터")
    
    # [원단 규격 계산기]
    st.subheader("📏 원단 규격 정밀 계산기")
    calc_mode = st.radio("계산 모드", ["⚖️ 무게 산출", "🔍 두께 역산"], horizontal=True)
    c1, c2, c3 = st.columns(3)
    with c1: v_width = st.number_input("비닐 폭 (mm)", value=630)
    with c2: v_length = st.number_input("원단 총 길이 (m)", value=1800)
    
    res_weight = 0
    if calc_mode == "⚖️ 무게 산출":
        with c3: v_thick = st.number_input("두께 (mm)", value=0.009, format="%.3f")
        res_weight = (v_width/1000) * v_length * 2 * 0.92 * v_thick
        st.info(f"💡 예상 무게: {res_weight:.2f} kg")
    else:
        with c3: v_weight_in = st.number_input("실제 무게 (kg)", value=13.8)
        res_thick = v_weight_in / ((v_width/1000) * v_length * 2 * 0.92)
        st.warning(f"💡 역산된 두께: {res_thick:.4f} mm")

    # [원재료 혼합 단가]
    st.divider()
    st.subheader("🧪 원재료 혼합 단가 계산")
    # (기존의 단가 계산 로직 실행)
    st.metric("최종 제조 원가", "데이터를 입력하세요")

# --- 메뉴 5: 시장 지표 분석 ---
elif menu == "📈 시장 지표 분석":
    st.title("📈 실시간 유가 및 환율 분석")
    if st.button("📊 최신 데이터 불러오기"):
        # (기존의 yfinance 및 그래프 시각화 로직 실행)
        st.success("지표 분석 완료")
