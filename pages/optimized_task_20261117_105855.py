import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta

# ==============================================================================
# 1. 시스템 환경 설정 및 전역 테마 스타일링
# ==============================================================================
st.set_page_config(
    page_title="요정비닐 통합 ERP 시스템",
    page_icon="🧚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ERP 전용 모던 CSS 스타일 적용
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 {
        margin: 0;
        font-size: 26px;
        font-weight: 700;
        color: #ffffff;
    }
    .main-header p {
        margin: 6px 0 0 0;
        font-size: 14px;
        opacity: 0.9;
    }
    
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-title {
        font-size: 13px;
        font-weight: 600;
        color: #64748b;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 24px;
        font-weight: 700;
        color: #0f172a;
    }
    .metric-delta {
        font-size: 12px;
        margin-top: 4px;
    }
    .metric-delta.plus { color: #10b981; }
    .metric-delta.minus { color: #ef4444; }

    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    
    div[data-testid="stSidebarNav"] {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. 세션 상태(Session State) 기반 데이터베이스 초기화
# ==============================================================================
def initialize_data():
    if "orders" not in st.session_state:
        st.session_state.orders = pd.DataFrame([
            {"주문번호": "ORD-20250501-01", "고객사": "(주)동아식품", "품목": "HDPE 롤백 (250*350)", "수량": 50000, "단가": 38, "수주일": "2025-05-01", "납기일": "2025-05-10", "진행상태": "생산중"},
            {"주문번호": "ORD-20250502-02", "고객사": "한솔패키징", "품목": "LDPE 지퍼백 (200*280)", "수량": 30000, "단가": 72, "수주일": "2025-05-02", "납기일": "2025-05-12", "진행상태": "주문접수"},
            {"주문번호": "ORD-20250503-03", "고객사": "모던상사", "품목": "PP 택배봉투 (350*450)", "수량": 80000, "단가": 95, "수주일": "2025-05-03", "납기일": "2025-05-08", "진행상태": "출하대기"},
            {"주문번호": "ORD-20250503-04", "고객사": "그린마켓", "품목": "생분해 비닐봉투 (대)", "수량": 20000, "단가": 120, "수주일": "2025-05-03", "납기일": "2025-05-15", "진행상태": "주문접수"},
            {"주문번호": "ORD-20250504-05", "고객사": "삼원화학", "품목": "OPP 테이프형 봉투", "수량": 15000, "단가": 45, "수주일": "2025-05-04", "납기일": "2025-05-07", "진행상태": "출하완료"}
        ])

    if "production" not in st.session_state:
        st.session_state.production = pd.DataFrame([
            {"작업지시번호": "WO-2405-01", "호기": "압출 1호기", "품목": "HDPE 필름 원단", "목표량(kg)": 1200, "실적량(kg)": 1180, "가동상태": "정상가동", "담당자": "김반장"},
            {"작업지시번호": "WO-2405-02", "호기": "인쇄 2호기", "품목": "동아식품 2도 인쇄", "목표량(m)": 8500, "실적량(m)": 8500, "가동상태": "작업완료", "담당자": "이주임"},
            {"작업지시번호": "WO-2405-03", "호기": "가공/실링 1호기", "품목": "롤백 열선 실링", "목표량(ea)": 50000, "실적량(ea)": 32000, "가동상태": "작업중", "담당자": "박기사"},
            {"작업지시번호": "WO-2405-04", "호기": "압출 2호기", "품목": "LDPE 투명 튜빙", "목표량(kg)": 900, "실적량(kg)": 0, "가동상태": "설비점검", "담당자": "최과장"}
        ])

    if "inventory" not in st.session_state:
        st.session_state.inventory = pd.DataFrame([
            {"코드": "RAW-HDPE-01", "품목명": "HDPE 수지 펠렛 (LG화학)", "분류": "원원자재", "규격": "25kg/포", "재고량": 14500, "단위": "kg", "안전재고": 5000, "위치": "원자재창고 A-1"},
            {"코드": "RAW-LDPE-02", "품목명": "LDPE 수지 펠렛 (한화솔루션)", "분류": "원원자재", "규격": "25kg/포", "재고량": 8200, "단위": "kg", "안전재고": 4000, "위치": "원자재창고 A-2"},
            {"코드": "INK-WH-01", "품목명": "그라비아 백색 잉크", "분류": "부자재", "규격": "18L/캔", "재고량": 14, "단위": "can", "안전재고": 10, "위치": "잉크보관실 B-3"},
            {"코드": "FIN-ROL-01", "품목명": "HDPE 롤백 규격품(완제품)", "분류": "완제품", "규격": "1박스(2000매)", "재고량": 180, "단위": "box", "안전재고": 50, "위치": "완제품창고 C-1"},
            {"코드": "FIN-ZIP-02", "품목명": "LDPE 투명지퍼백(완제품)", "분류": "완제품", "규격": "1박스(1000매)", "재고량": 35, "단위": "box", "안전재고": 50, "위치": "완제품창고 C-2"}
        ])

    if "partners" not in st.session_state:
        st.session_state.partners = pd.DataFrame([
            {"거래처코드": "CUST-001", "거래처명": "(주)동아식품", "구분": "매출처", "사업자번호": "120-81-45678", "대표자": "김동아", "담당자연락처": "010-2345-6789"},
            {"거래처코드": "CUST-002", "거래처명": "한솔패키징", "구분": "매출처", "사업자번호": "214-85-98765", "대표자": "이한솔", "담당자연락처": "010-9876-5432"},
            {"거래처코드": "SUPP-001", "거래처명": "LG화학 원료유통", "구분": "매입처", "사업자번호": "105-86-12345", "대표자": "신원료", "담당자연락처": "02-3773-1114"},
            {"거래처코드": "SUPP-002", "거래처명": "대한특수잉크", "구분": "매입처", "사업자번호": "131-82-44332", "대표자": "정컬러", "담당자연락처": "031-492-8877"}
        ])

initialize_data()

# ==============================================================================
# 3. 서브 모듈 렌더링 함수군 (develop/pages 모듈 통합)
# ==============================================================================

def render_dashboard():
    st.markdown("### 📊 통합 경영 및 제조 현황 관제")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">금월 총 수주금액</div>
            <div class="metric-value">₩ 14,845,000</div>
            <div class="metric-delta plus">▲ 전월 대비 12.4% 증가</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">진행중 작업 건수</div>
            <div class="metric-value">6 건</div>
            <div class="metric-delta plus">압출 2건 / 실링 4건</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">금일 출하 예정</div>
            <div class="metric-value">1 건</div>
            <div class="metric-delta minus">모던상사 (80,000매)</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-title">안전재고 미달 품목</div>
            <div class="metric-value">1 건</div>
            <div class="metric-delta minus">LDPE 투명지퍼백 주의</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    
    col_left, col_right = st.columns([7, 5])
    with col_left:
        st.subheader("📈 주간 일자별 생산량 추이 (kg)")
        trend_data = pd.DataFrame({
            "일자": ["05-01", "05-02", "05-03", "05-04", "05-05", "05-06", "05-07"],
            "HDPE": [1100, 1350, 1200, 1400, 950, 1500, 1300],
            "LDPE": [800, 920, 850, 900, 780, 1100, 1050],
            "PP": [400, 450, 500, 520, 300, 600, 550]
        })
        st.line_chart(trend_data.set_index("일자"))
        
    with col_right:
        st.subheader("⚡ 설비 가동 상태 요약")
        status_counts = st.session_state.production["가동상태"].value_counts().reset_index()
        status_counts.columns = ["상태", "호기수"]
        st.bar_chart(status_counts.set_index("상태"))
        
    st.subheader("📋 금일 주요 공정 진행 타임라인")
    st.dataframe(
        st.session_state.production,
        use_container_width=True,
        hide_index=True
    )

def render_orders():
    st.markdown("### 📝 수주 및 견적 관리")
    
    tab1, tab2 = st.tabs(["수주 목록 및 진행 현황", "신규 수주 등록"])
    
    with tab1:
        col_f1, col_f2 = st.columns([3, 1])
        with col_f1:
            search_kw = st.text_input("고객사 또는 품목명 검색", placeholder="예: 동아식품, 지퍼백")
        with col_f2:
            status_filter = st.selectbox("진행상태 필터", ["전체", "주문접수", "생산중", "출하대기", "출하완료"])
        
        filtered = st.session_state.orders.copy()
        if search_kw:
            filtered = filtered[filtered["고객사"].str.contains(search_kw) | filtered["품목"].str.contains(search_kw)]
        if status_filter != "전체":
            filtered = filtered[filtered["진행상태"] == status_filter]
            
        filtered["수주총액(원)"] = filtered["수량"] * filtered["단가"]
        st.dataframe(filtered, use_container_width=True, hide_index=True)
        
    with tab2:
        st.write("##### 신규 수주 내역 입력")
        with st.form("new_order_form", clear_on_submit=True):
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                cust = st.selectbox("발주처(고객사)", st.session_state.partners[st.session_state.partners["구분"]=="매출처"]["거래처명"].tolist())
                item = st.text_input("제품명(사양)", placeholder="예: HDPE 쇼핑백 인쇄 300*400")
                qty = st.number_input("수량(EA)", min_value=100, step=1000, value=10000)
            with f_col2:
                price = st.number_input("납품단가(원)", min_value=1, step=1, value=50)
                order_date = st.date_input("수주일자", date.today())
                due_date = st.date_input("납기일자", date.today() + timedelta(days=7))
            
            submitted = st.form_submit_button("신규 수주 접수 저장", use_container_width=True)
            if submitted:
                new_idx = len(st.session_state.orders) + 1
                new_row = {
                    "주문번호": f"ORD-{order_date.strftime('%Y%m%d')}-{new_idx:02d}",
                    "고객사": cust,
                    "품목": item,
                    "수량": int(qty),
                    "단가": int(price),
                    "수주일": str(order_date),
                    "납기일": str(due_date),
                    "진행상태": "주문접수"
                }
                st.session_state.orders = pd.concat([st.session_state.orders, pd.DataFrame([new_row])], ignore_index=True)
                st.success("새로운 수주건이 정상 등록되었습니다.")
                st.rerun()

def render_production():
    st.markdown("### ⚙️ 생산 공정 및 작업지시 관리")
    
    col1, col2 = st.columns([8, 4])
    with col1:
        st.subheader("호기별 작업 실적 현황")
        st.dataframe(st.session_state.production, use_container_width=True, hide_index=True)
    
    with col2:
        st.subheader("실시간 공정 등록/수정")
        with st.form("prod_form"):
            wo = st.selectbox("작업지시 선택", st.session_state.production["작업지시번호"].tolist())
            add_qty = st.number_input("추가 실적 수량", min_value=0, value=100)
            new_status = st.selectbox("설비 상태 변경", ["정상가동", "작업중", "작업완료", "설비점검", "일시대기"])
            
            if st.form_submit_button("작업실적 반영"):
                row_idx = st.session_state.production[st.session_state.production["작업지시번호"] == wo].index[0]
                target_col = "실적량(kg)" if "kg" in st.session_state.production.columns[3] else "실적량(ea)"
                st.session_state.production.loc[row_idx, "가동상태"] = new_status
                st.session_state.production.loc[row_idx, target_col] += add_qty
                st.success(f"{wo} 실적이 성공적으로 갱신되었습니다.")
                st.rerun()

def render_inventory():
    st.markdown("### 📦 원자재 및 완제품 재고 현황")
    
    inv_df = st.session_state.inventory.copy()
    inv_df["재고상태"] = inv_df.apply(lambda r: "🚨 재고부족" if r["재고량"] < r["안전재고"] else "✅ 정상", axis=1)
    
    st.dataframe(inv_df, use_container_width=True, hide_index=True)
    
    st.divider()
    st.write("##### 🔄 원단 및 부자재 입/출고 등록")
    i_col1, i_col2, i_col3, i_col4 = st.columns(4)
    with i_col1:
        sel_item = st.selectbox("품목 선택", inv_df["품목명"].tolist())
    with i_col2:
        in_out = st.radio("구분", ["입고", "출고"], horizontal=True)
    with i_col3:
        io_qty = st.number_input("입출고량", min_value=1, value=100)
    with i_col4:
        st.write("")
        st.write("")
        if st.button("수불 저장", use_container_width=True):
            r_idx = st.session_state.inventory[st.session_state.inventory["품목명"] == sel_item].index[0]
            curr_qty = st.session_state.inventory.loc[r_idx, "재고량"]
            if in_out == "출고" and curr_qty < io_qty:
                st.error("현재고보다 많은 수량을 출고할 수 없습니다.")
            else:
                new_qty = curr_qty + io_qty if in_out == "입고" else curr_qty - io_qty
                st.session_state.inventory.loc[r_idx, "재고량"] = new_qty
                st.success(f"처리 완료되었습니다. 변경 후 재고: {new_qty:,}")
                st.rerun()

def render_logistics():
    st.markdown("### 🚚 출하 및 물류 배송 관리")
    
    ship_targets = st.session_state.orders[st.session_state.orders["진행상태"].isin(["출하대기", "출하완료"])]
    st.write("##### 출하 대기 및 완료 목록")
    st.dataframe(ship_targets, use_container_width=True, hide_index=True)
    
    st.divider()
    st.write("##### 송장 발행 및 출하 승인")
    col1, col2 = st.columns(2)
    with col1:
        pending_orders = st.session_state.orders[st.session_state.orders["진행상태"] == "출하대기"]
        if not pending_orders.empty:
            sel_ord = st.selectbox("출하 대기 전표", pending_orders["주문번호"].tolist())
            courier = st.selectbox("운송수단 / 택배사", ["자사 화물트럭(1톤)", "지입 화물트럭(5톤)", "경동택배", "로젠택배"])
            driver_info = st.text_input("기사명/연락처", "홍길동 기사 (010-5555-1234)")
            if st.button("출하 승인 및 거래명세서 발행"):
                idx = st.session_state.orders[st.session_state.orders["주문번호"] == sel_ord].index[0]
                st.session_state.orders.loc[idx, "진행상태"] = "출하완료"
                st.success(f"{sel_ord} 출하 처리가 완료되었습니다.")
                st.rerun()
        else:
            st.info("현재 대기중인 출하 건이 없습니다.")

def render_settlement():
    st.markdown("### 💰 정산 및 매출 분석")
    
    orders = st.session_state.orders.copy()
    orders["총액"] = orders["수량"] * orders["단가"]
    
    total_sales = orders[orders["진행상태"] == "출하완료"]["총액"].sum()
    pending_sales = orders[orders["진행상태"] != "출하완료"]["총액"].sum()
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">완료된 확정 매출(공급가액)</div>
            <div class="metric-value">₩ {total_sales:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">미출하 진행중 매출 예정액</div>
            <div class="metric-value">₩ {pending_sales:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.write("")
    st.subheader("거래처별 정산 명세")
    partner_sales = orders.groupby("고객사")["총액"].sum().reset_index()
    partner_sales.columns = ["거래처명", "주문 합계액(원)"]
    st.dataframe(partner_sales, use_container_width=True, hide_index=True)

def render_partners():
    st.markdown("### 🏢 기준정보 (거래처 및 품목 원장)")
    
    st.write("##### 등록 거래처 마스터")
    st.dataframe(st.session_state.partners, use_container_width=True, hide_index=True)
    
    with st.expander("➕ 신규 거래처 추가"):
        with st.form("partner_add_form"):
            c1, c2 = st.columns(2)
            with c1:
                p_code = st.text_input("거래처 코드", value=f"CUST-{len(st.session_state.partners)+1:03d}")
                p_name = st.text_input("거래처 상호명")
                p_type = st.selectbox("구분", ["매출처", "매입처", "기타"])
            with c2:
                p_biz_no = st.text_input("사업자등록번호", placeholder="000-00-00000")
                p_rep = st.text_input("대표자명")
                p_tel = st.text_input("담당자 연락처")
                
            if st.form_submit_button("거래처 등록 저장"):
                if p_name:
                    new_p = {
                        "거래처코드": p_code,
                        "거래처명": p_name,
                        "구분": p_type,
                        "사업자번호": p_biz_no,
                        "대표자": p_rep,
                        "담당자연락처": p_tel
                    }
                    st.session_state.partners = pd.concat([st.session_state.partners, pd.DataFrame([new_p])], ignore_index=True)
                    st.success(f"{p_name} 거래처가 등록되었습니다.")
                    st.rerun()
                else:
                    st.error("거래처 상호명을 입력해주세요.")

# ==============================================================================
# 4. 사이드바 네비게이션 및 메인 라우터
# ==============================================================================
def main():
    # 사이드바 레이아웃
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 10px 0;">
            <h2 style="margin: 0; color: #1e3a8a;">🧚 요정비닐 ERP</h2>
            <span style="font-size: 12px; color: #64748b;">제조·압출·가공 통합운영체계</span>
        </div>
        """, unsafe_allow_html=True)
        st.divider()
        
        menu = st.radio(
            "시스템 메뉴",
            [
                "📊 종합 대시보드",
                "📝 수주/영업 관리",
                "⚙️ 생산/공정 관리",
                "📦 자재/재고 관리",
                "🚚 출하/물류 관리",
                "💰 매출/정산 관리",
                "🏢 기준정보 관리"
            ],
            index=0
        )
        
        st.divider()
        st.markdown("""
        <div style="font-size: 12px; color: #94a3b8; line-height: 1.6;">
            <b>접속 계정:</b> 관리자 (공장운영팀)<br>
            <b>최종 동기화:</b> """ + datetime.now().strftime("%Y-%m-%d %H:%M") + """<br>
            <b>시스템 상태:</b> <span style="color:#10b981;">● 온라인 (정상)</span>
        </div>
        """, unsafe_allow_html=True)

    # 상단 공통 헤더
    st.markdown("""
    <div class="main-header">
        <h1>요정비닐 제조실행 및 전사자원관리 시스템 (ERP)</h1>
        <p>압출, 인쇄, 가공 공정 추적부터 재고 수불 및 출하 정산까지 단일 환경에서 제어합니다.</p>
    </div>
    """, unsafe_allow_html=True)

    # 메뉴 라우팅
    if menu == "📊 종합 대시보드":
        render_dashboard()
    elif menu == "📝 수주/영업 관리":
        render_orders()
    elif menu == "⚙️ 생산/공정 관리":
        render_production()
    elif menu == "📦 자재/재고 관리":
        render_inventory()
    elif menu == "🚚 출하/물류 관리":
        render_logistics()
    elif menu == "💰 매출/정산 관리":
        render_settlement()
    elif menu == "🏢 기준정보 관리":
        render_partners()

if __name__ == "__main__":
    main()