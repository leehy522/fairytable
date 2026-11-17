import streamlit as st
import pandas as pd
import datetime

# --- 페이지 설정 ---
st.set_page_config(
    page_title="요정비닐 ERP 시스템",
    page_icon="🧚‍♀️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- 커스텀 스타일 (엔터프라이즈 ERP 테마) ---
st.markdown("""
<style>
    /* 글로벌 폰트 및 배경 미세 조정 */
    .stApp {
        background-color: #f8fafc;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    
    /* 사이드바 스타일링 */
    section[data-testid="stSidebar"] {
        background-color: #1e293b;
        color: #f1f5f9;
        border-right: 1px solid #334155;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] label {
        color: #f1f5f9 !important;
    }
    
    /* 사이드바 메뉴 그룹 배지 */
    .menu-header {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94a3b8 !important;
        margin-top: 1.25rem;
        margin-bottom: 0.5rem;
        padding-left: 0.25rem;
    }
    
    /* 지표 카드 */
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    .metric-card .title {
        font-size: 0.875rem;
        color: #64748b;
        margin-bottom: 6px;
    }
    .metric-card .value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0f172a;
    }
    .metric-card .delta {
        font-size: 0.8rem;
        font-weight: 600;
    }
    .delta-up { color: #10b981; }
    .delta-down { color: #ef4444; }

    /* 헤더 배너 */
    .system-badge {
        display: inline-block;
        padding: 3px 8px;
        font-size: 0.75rem;
        font-weight: 600;
        border-radius: 4px;
        background-color: #e0e7ff;
        color: #4338ca;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# --- 사이드바: 브랜드 및 구조화된 메뉴 체계 ---
with st.sidebar:
    st.markdown("""
        <div style="padding: 10px 0 20px 0; border-bottom: 1px solid #334155;">
            <div style="font-size: 1.3rem; font-weight: 800; color: #38bdf8; display: flex; align-items: center; gap: 8px;">
                <span>🧚‍♀️ 요정비닐</span>
            </div>
            <div style="font-size: 0.75rem; color: #94a3b8; margin-top: 4px;">통합 생산·영업 관리 시스템 v2.4</div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="menu-header">시스템 메뉴</div>', unsafe_allow_html=True)
    
    # 2단계 계층형 구조: 대메뉴(카테고리) 및 소메뉴
    MENU_STRUCTURE = {
        "📊 종합 대시보드": ["실시간 종합 현황판"],
        "💼 영업 및 주문 관리": ["수주/주문 접수 내역", "신규 발주서 등록", "거래처 원장"],
        "🏭 생산 및 공정 관리": ["생산 계획 및 지시", "공정별 모니터링 (압출/인쇄/가공)", "설비 가동 현황"],
        "📦 재고 및 원자재": ["원자재 수불 현황 (PE/PP)", "완제품 창고 재고", "입출고 전표 관리"],
        "🚚 물류 및 출하": ["출하 지시/배송 관리", "차량별 배차 현황"],
        "📈 경영 및 통계": ["월별/품목별 매출 통계", "생산 수율 및 불량률 분석"],
        "⚙️ 시스템 환경설정": ["기본 코드 관리", "사용자 및 권한 설정"]
    }
    
    selected_category = st.selectbox(
        "카테고리 선택",
        options=list(MENU_STRUCTURE.keys()),
        index=0,
        label_visibility="collapsed"
    )
    
    sub_menu_options = MENU_STRUCTURE[selected_category]
    selected_menu = st.radio(
        "상세 화면 선택",
        options=sub_menu_options,
        index=0,
        label_visibility="collapsed"
    )
    
    # 사이드바 하단 사용자 및 작업자 정보
    st.markdown("---")
    st.markdown("""
        <div style="font-size: 0.75rem; color: #94a3b8; line-height: 1.5;">
            <div>👤 <b>접속자:</b> 김생산 관리자 (생산본부)</div>
            <div>🟢 <b>MES 연동:</b> 정상 작동 중</div>
            <div>⏱ <b>동기화:</b> {}</div>
        </div>
    """.format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M")), unsafe_allow_html=True)

# --- 상단 타이틀 영역 ---
st.markdown('<div class="system-badge">ERP Navigation Console</div>', unsafe_allow_html=True)
st.title(f"{selected_menu}")
st.caption(f"선택 분류: {selected_category} > {selected_menu}")
st.divider()

# --- 화면별 기능 뷰 구현 ---
if selected_menu == "실시간 종합 현황판":
    # 상단 요약 KPI 카드
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("""
            <div class="metric-card">
                <div class="title">금일 접수 주문</div>
                <div class="value">48 건</div>
                <div class="delta delta-up">▲ 12% (전일 대비)</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
            <div class="metric-card">
                <div class="title">생산 라인 가동률</div>
                <div class="value">94.2%</div>
                <div class="delta delta-up">▲ 1.8% (목표 대비)</div>
            </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
            <div class="metric-card">
                <div class="title">원자재 안전재고율 (LDPE)</div>
                <div class="value">118%</div>
                <div class="delta delta-up">● 적정 재고 유지</div>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown("""
            <div class="metric-card">
                <div class="title">금일 출하 대기</div>
                <div class="value">15 건</div>
                <div class="delta delta-down">▼ 3건 지연 위험</div>
            </div>
        """, unsafe_allow_html=True)

    # 탭 구성으로 복잡한 정보 그룹화
    tab1, tab2 = st.tabs(["🔥 주요 생산 라인 상태", "📋 최근 주문 접수 목록"])
    
    with tab1:
        st.subheader("압출 및 가공 1·2공장 라인 현황")
        lines_data = pd.DataFrame([
            {"호기": "압출 1호기", "규격": "HDPE 0.03*450*600", "작업상태": "가동중", "진행률": "82%", "담당자": "박진철"},
            {"호기": "압출 2호기", "규격": "LDPE 0.05*300*400", "작업상태": "가동중", "진행률": "45%", "담당자": "이동수"},
            {"호기": "인쇄 1호기 (4도)", "규격": "요정마트 2도 인쇄롤", "작업상태": "대기중", "진행률": "-", "담당자": "최영민"},
            {"호기": "가공/실링 1호기", "규격": "배달용 평판비닐", "작업상태": "가동중", "진행률": "95%", "담당자": "강호준"},
            {"호기": "가공/실링 2호기", "규격": "손잡이형 M방 봉투", "작업상태": "정비중", "진행률": "0%", "담당자": "외부AS"},
        ])
        st.dataframe(lines_data, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("실시간 수주 진행 내역")
        orders_data = pd.DataFrame([
            {"주문번호": "ORD-20250520-01", "거래처명": "(주)싱싱유통", "제품명": "HDPE 롤백 30*40", "수량(Box)": 200, "납기일": "2025-05-24", "상태": "생산 배정"},
            {"주문번호": "ORD-20250520-02", "거래처명": "요정물류상사", "제품명": "택배용 LDPE 2중봉투", "수량(Box)": 500, "납기일": "2025-05-23", "상태": "압출 완료"},
            {"주문번호": "ORD-20250520-03", "거래처명": "삼정식품", "제품명": "진공포장용 롤필름", "수량(Box)": 80, "납기일": "2025-05-22", "상태": "출하 대기"},
            {"주문번호": "ORD-20250520-04", "거래처명": "에코그린팩토리", "제품명": "생분해성 쇼핑백", "수량(Box)": 150, "납기일": "2025-05-26", "상태": "원료 준비"},
        ])
        st.dataframe(orders_data, use_container_width=True, hide_index=True)

elif selected_menu == "수주/주문 접수 내역":
    col_filter1, col_filter2, col_filter3 = st.columns([2, 2, 3])
    with col_filter1:
        st.date_input("조회 기간 시작", datetime.date.today() - datetime.timedelta(days=7))
    with col_filter2:
        st.date_input("조회 기간 종료", datetime.date.today())
    with col_filter3:
        st.text_input("거래처명 또는 제품명 검색", placeholder="예: 싱싱유통, HDPE")

    st.markdown("#### 검색 결과")
    order_master = pd.DataFrame([
        {"주문번호": "ORD-20250519-01", "일자": "2025-05-19", "거래처": "(주)대호포장", "품목": "HDPE 평판 0.02t", "수량": "50,000매", "금액": "1,850,000원", "결제상태": "외상매출", "공정상태": "가공중"},
        {"주문번호": "ORD-20250519-02", "일자": "2025-05-19", "거래처": "온누리마켓", "품목": "쇼핑백 중형 2도", "수량": "20,000매", "금액": "980,000원", "결제상태": "입금완료", "공정상태": "출하완료"},
        {"주문번호": "ORD-20250518-01", "일자": "2025-05-18", "거래처": "케이팩트리", "품목": "공업용 스트레치필름", "수량": "300롤", "금액": "3,600,000원", "결제상태": "계리의뢰", "공정상태": "출하완료"},
    ])
    st.dataframe(order_master, use_container_width=True, hide_index=True)

elif selected_menu == "신규 발주서 등록":
    st.subheader("📝 주문서 접수 양식")
    with st.form(key="new_order_form"):
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            st.selectbox("거래처 선택", ["(주)싱싱유통", "요정물류상사", "삼정식품", "신규 거래처 직접 입력"])
            st.selectbox("비닐 재질 선택", ["HDPE (고밀도 폴리에틸렌)", "LDPE (저밀도 폴리에틸렌)", "PP (폴리프로필렌)", "생분해 수지 (PLA)"])
            st.text_input("상세 규격 (두께T × 가로mm × 세로mm)", placeholder="예: 0.04 × 400 × 600")
        with f_col2:
            st.number_input("주문 수량 (단위: 매 또는 롤)", min_value=100, step=100, value=1000)
            st.date_input("희망 납기일", datetime.date.today() + datetime.timedelta(days=5))
            st.selectbox("인쇄 여부", ["무지 (인쇄 없음)", "1도 인쇄", "2도 인쇄", "4도 컬러 인쇄"])
        
        st.text_area("특이사항 및 가공 지침 (예: 타공 처리, M방 접지 여부)", placeholder="상세 요청사항을 입력하세요.")
        submit_btn = st.form_submit_button("주문 접수 및 생산지시 전송", type="primary", use_container_width=True)
        if submit_btn:
            st.success("✅ 주문이 정상 접수되었습니다. 생산계획팀에 자동으로 전송됩니다.")

elif selected_menu == "원자재 수불 현황 (PE/PP)":
    st.subheader("원자재(레진/펠릿) 및 첨가제 재고 내역")
    raw_materials = pd.DataFrame([
        {"품목코드": "RM-HD-001", "품목명": "HDPE 펠릿 (기본형)", "현재고": "18,400 kg", "안전재고": "8,000 kg", "단가(kg)": "1,450원", "재고상태": "안전"},
        {"품목코드": "RM-LD-002", "품목명": "LDPE 고투명 펠릿", "현재고": "6,200 kg", "안전재고": "7,000 kg", "단가(kg)": "1,620원", "재고상태": "주의 (발주요망)"},
        {"품목코드": "RM-MB-001", "품목명": "화이트 마스터배치", "현재고": "1,250 kg", "안전재고": "500 kg", "단가(kg)": "2,100원", "재고상태": "안전"},
        {"품목코드": "RM-MB-002", "품목명": "블랙/차광 마스터배치", "현재고": "420 kg", "안전재고": "400 kg", "단가(kg)": "2,300원", "재고상태": "적정"},
    ])
    st.dataframe(raw_materials, use_container_width=True, hide_index=True)
    
    st.info("💡 안전재고 수량 미달 품목은 원자재 수급 알림이 구매팀으로 자동 발송됩니다.")

else:
    # 기타 메뉴 기본 안내 템플릿
    st.info(f"선택하신 메뉴: **[{selected_menu}]**")
    st.write("해당 메뉴의 데이터 스트림 및 MES 연결 프로토콜이 대기 중입니다.")
    st.caption("기존 시스템의 비즈니스 로직과 데이터베이스 연동 함수를 이 인터페이스에 바로 연결하여 사용할 수 있습니다.")