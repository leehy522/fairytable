import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import datetime
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="요정비닐 ERP v2.5",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.2rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        text-align: center;
    }
    
    .status-badge {
        display: inline-block;
        padding: 0.25rem 0.6rem;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 600;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        border-radius: 8px;
        font-weight: 600;
        padding: 0 20px;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

class DatabaseManager:
    def __init__(self, db_name="fairy_vinyl_erp.db"):
        self.db_name = db_name
        self.init_db()

    def get_connection(self):
        return sqlite3.connect(self.db_name)

    def init_db(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                order_id TEXT PRIMARY KEY,
                client_name TEXT NOT NULL,
                material TEXT NOT NULL,
                thickness_t REAL NOT NULL,
                width_w REAL NOT NULL,
                length_l REAL NOT NULL,
                fold_m REAL DEFAULT 0,
                printing TEXT DEFAULT '무인쇄',
                order_qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                total_amount REAL NOT NULL,
                order_date TEXT NOT NULL,
                due_date TEXT NOT NULL,
                status TEXT NOT NULL
            )
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS production_schedule (
                plan_id TEXT PRIMARY KEY,
                order_id TEXT NOT NULL,
                line_no TEXT NOT NULL,
                process_step TEXT NOT NULL,
                target_qty INTEGER NOT NULL,
                completed_qty INTEGER DEFAULT 0,
                worker TEXT NOT NULL,
                start_date TEXT,
                end_date TEXT,
                status TEXT NOT NULL,
                FOREIGN KEY(order_id) REFERENCES orders(order_id)
            )
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS inventory (
                item_code TEXT PRIMARY KEY,
                item_name TEXT NOT NULL,
                category TEXT NOT NULL,
                spec TEXT,
                current_stock REAL NOT NULL,
                safety_stock REAL NOT NULL,
                unit TEXT NOT NULL,
                unit_cost REAL NOT NULL,
                last_updated TEXT NOT NULL
            )
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                client_id TEXT PRIMARY KEY,
                client_name TEXT NOT NULL,
                contact TEXT NOT NULL,
                phone TEXT NOT NULL,
                business_number TEXT,
                address TEXT
            )
            """)
            conn.commit()
            self._seed_default_data(conn)

    def _seed_default_data(self, conn):
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM inventory")
        if cursor.fetchone()[0] == 0:
            sample_inv = [
                ("RAW-LDPE-01", "LDPE 수지 (원료)", "원자재", "LG화학 SP980", 4200.0, 1500.0, "kg", 1850, "2025-02-18"),
                ("RAW-HDPE-01", "HDPE 수지 (원료)", "원자재", "롯데케미칼 7000F", 5800.0, 2000.0, "kg", 1750, "2025-02-18"),
                ("RAW-LLDPE-01", "LLDPE 원료", "원자재", "한화솔루션 3304", 1200.0, 1500.0, "kg", 1900, "2025-02-19"),
                ("MB-WHT-01", "화이트 마스터배치", "부자재", "백색 농축착색제", 450.0, 200.0, "kg", 4500, "2025-02-15"),
                ("MB-BLK-01", "블랙 마스터배치", "부자재", "흑색 농축착색제", 180.0, 200.0, "kg", 4200, "2025-02-15"),
                ("INK-BLU-01", "그라비아 잉크 청색", "부자재", "친환경 속건형", 85.0, 50.0, "kg", 8500, "2025-02-10"),
                ("FIN-PB-001", "택배용 LDPE 비닐봉투", "완제품", "0.06T * 350W * 450L", 12500.0, 5000.0, "매", 65, "2025-02-20"),
                ("FIN-HD-002", "마트용 HDPE 롤백", "완제품", "0.02T * 250W * 350L", 320.0, 100.0, "롤", 4500, "2025-02-20")
            ]
            cursor.executemany("INSERT INTO inventory VALUES (?,?,?,?,?,?,?,?,?)", sample_inv)

        cursor.execute("SELECT COUNT(*) FROM clients")
        if cursor.fetchone()[0] == 0:
            sample_clients = [
                ("CLI-001", "동서포장물류", "김부장", "010-3344-9988", "120-81-99231", "경기 파주시 탄현면"),
                ("CLI-002", "그린마켓 유통", "이팀장", "010-8877-2211", "214-86-44512", "서울 성동구 성수동"),
                ("CLI-003", "신세계팩", "박이사", "010-4455-6677", "131-82-66321", "충남 천안시 서북구"),
                ("CLI-004", "요정패키지 직영점", "최과장", "010-1122-3344", "105-88-00192", "경기 고양시 일산동구")
            ]
            cursor.executemany("INSERT INTO clients VALUES (?,?,?,?,?,?)", sample_clients)

        cursor.execute("SELECT COUNT(*) FROM orders")
        if cursor.fetchone()[0] == 0:
            sample_orders = [
                ("ORD-202502-001", "동서포장물류", "LDPE", 0.05, 300, 400, 0, "1도 인쇄", 50000, 48.0, 2400000.0, "2025-02-15", "2025-02-23", "생산중"),
                ("ORD-202502-002", "그린마켓 유통", "HDPE", 0.03, 250, 350, 40, "무인쇄", 100000, 22.0, 2200000.0, "2025-02-16", "2025-02-25", "수주접수"),
                ("ORD-202502-003", "신세계팩", "OPP", 0.04, 200, 300, 0, "2도 양면", 30000, 65.0, 1950000.0, "2025-02-17", "2025-02-22", "가공완료"),
                ("ORD-202502-004", "요정패키지 직영점", "LLDPE", 0.06, 500, 600, 50, "1도 전면", 20000, 110.0, 2200000.0, "2025-02-18", "2025-02-28", "압출대기")
            ]
            cursor.executemany("INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", sample_orders)

        cursor.execute("SELECT COUNT(*) FROM production_schedule")
        if cursor.fetchone()[0] == 0:
            sample_sched = [
                ("PLN-001", "ORD-202502-001", "압출 1호기", "1단계 압출", 50000, 50000, "강기사", "2025-02-16", "2025-02-17", "완료"),
                ("PLN-002", "ORD-202502-001", "인쇄 2호기", "2단계 인쇄", 50000, 35000, "윤반장", "2025-02-18", "2025-02-19", "진행중"),
                ("PLN-003", "ORD-202502-003", "실링가공기 A", "3단계 가공/실링", 30000, 30000, "송조장", "2025-02-19", "2025-02-20", "완료")
            ]
            cursor.executemany("INSERT INTO production_schedule VALUES (?,?,?,?,?,?,?,?,?,?)", sample_sched)
            conn.commit()

db = DatabaseManager()

def fetch_data(query, params=()):
    with db.get_connection() as conn:
        return pd.read_sql_query(query, conn, params=params)

def execute_query(query, params=()):
    with db.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()

def calculate_vinyl_weight(t_mm, w_mm, l_mm, qty, density=0.92):
    single_weight_g = (t_mm * (w_mm / 10) * (l_mm / 10) * density * 2) / 100
    total_kg = (single_weight_g * qty) / 1000.0
    return round(single_weight_g, 3), round(total_kg, 2)

st.markdown("""
<div class="main-header">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <h1 style="margin:0; font-size:2rem; font-weight:700;">🧚 요정비닐 통합 ERP 시스템</h1>
            <p style="margin:0.3rem 0 0 0; opacity:0.85; font-size:0.95rem;">공정 압출 · 인쇄 · 가공 재고실시간 트래킹 솔루션</p>
        </div>
        <div style="text-align:right;">
            <span style="background:rgba(255,255,255,0.2); padding:6px 14px; border-radius:20px; font-size:0.85rem; font-weight:600;">
                🟢 표준 공정 가동중
            </span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://images.unsplash.com/photo-1581091226825-a6a2a5aee158?auto=format&fit=crop&w=400&q=80", use_container_width=True)
    st.subheader("📌 ERP 제어 패널")
    menu = st.radio(
        "메뉴 바로가기",
        ["📊 종합 현황 대시보드", "📋 수주 및 주문 관리", "⚙️ 공정 및 생산 관리", "📦 원자재 및 완제품 재고", "👥 거래처 및 출하 관리", "🧮 비닐 중량/원가 계산기"],
        index=0
    )
    st.divider()
    st.caption("요정비닐 공업사 (Fairy Vinyl Co.)")
    st.caption("시스템 버전: v2.5.0-LTS (Build 202502)")

if menu == "📊 종합 현황 대시보드":
    df_orders = fetch_data("SELECT * FROM orders")
    df_inv = fetch_data("SELECT * FROM inventory")
    df_prod = fetch_data("SELECT * FROM production_schedule")

    total_sales = df_orders['total_amount'].sum()
    active_orders = len(df_orders[df_orders['status'].isin(['수주접수', '압출대기', '생산중'])])
    alert_inv = len(df_inv[df_inv['current_stock'] <= df_inv['safety_stock']])
    running_lines = len(df_prod[df_prod['status'] == '진행중'])

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("총 수주 누적액", f"{total_sales:,.0f} 원", "+12.4%")
    with col2:
        st.metric("진행 중인 발주건", f"{active_orders} 건", f"전체 {len(df_orders)}건 중")
    with col3:
        st.metric("안전재고 미달 품목", f"{alert_inv} 품목", "-2건" if alert_inv > 0 else "정상", delta_color="inverse")
    with col4:
        st.metric("가동 라인", f"{running_lines} 개소", "정상 가동")

    st.write(" ")
    g_col1, g_col2 = st.columns([6, 4])

    with g_col1:
        st.subheader("📈 원단 재질별 수주 수량 현황")
        if not df_orders.empty:
            mat_summary = df_orders.groupby('material')['order_qty'].sum().reset_index()
            fig = px.bar(
                mat_summary,
                x='material',
                y='order_qty',
                color='material',
                labels={'material': '원료 재질', 'order_qty': '총 주문수량 (EA/M)'},
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=320)
            st.plotly_chart(fig, use_container_width=True)

    with g_col2:
        st.subheader("🍩 주문 진행 상태 비율")
        if not df_orders.empty:
            status_summary = df_orders['status'].value_counts().reset_index()
            status_summary.columns = ['status', 'count']
            fig_pie = px.pie(
                status_summary,
                names='status',
                values='count',
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Safe
            )
            fig_pie.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=320)
            st.plotly_chart(fig_pie, use_container_width=True)

    st.subheader("🚨 원자재 재고 안전선 경보")
    low_stocks = df_inv[df_inv['current_stock'] <= df_inv['safety_stock']]
    if not low_stocks.empty:
        st.warning(f"현재 {len(low_stocks)}개 품목이 안전 재고량 이하입니다. 신속한 추가 발주가 필요합니다.")
        st.dataframe(low_stocks[['item_code', 'item_name', 'category', 'current_stock', 'safety_stock', 'unit']], use_container_width=True)
    else:
        st.success("모든 원자재 및 부자재가 안전재고 수량을 충족하고 있습니다.")

elif menu == "📋 수주 및 주문 관리":
    st.subheader("📝 신규 수주 등록 및 내역")
    tab_list, tab_new = st.tabs(["주문 관리 대장", "새 수주서 발행"])

    with tab_new:
        with st.form("new_order_form", clear_on_submit=True):
            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                client_options = fetch_data("SELECT client_name FROM clients")['client_name'].tolist()
                client = st.selectbox("발주처", client_options if client_options else ["(주)신규거래처"])
                mat = st.selectbox("원료 재질", ["LDPE", "HDPE", "LLDPE", "PP", "OPP"])
                printing = st.selectbox("인쇄 사양", ["무인쇄", "동판 1도 단면", "동판 1도 양면", "동판 2도 단면", "동판 2도 양면", "4도 칼라"])
            
            with f_col2:
                thickness = st.number_input("두께 (T, mm 단위)", min_value=0.010, max_value=0.200, value=0.050, step=0.005, format="%.3f")
                width = st.number_input("가로 폭 (W, mm)", min_value=50.0, max_value=2000.0, value=300.0, step=10.0)
                length = st.number_input("세로 길이 (L, mm)", min_value=50.0, max_value=3000.0, value=450.0, step=10.0)
                fold = st.number_input("M접이 가공폭 (M, mm - 단면 기준)", min_value=0.0, max_value=300.0, value=0.0, step=5.0)

            with f_col3:
                qty = st.number_input("발주 수량 (매)", min_value=500, max_value=5000000, value=10000, step=1000)
                unit_p = st.number_input("납품 단가 (원/매)", min_value=1.0, max_value=10000.0, value=45.0, step=5.0)
                due = st.date_input("납기 기한", datetime.date.today() + datetime.timedelta(days=7))

            calc_single_g, calc_tot_kg = calculate_vinyl_weight(thickness, width + (fold * 2), length, qty)
            total_price = qty * unit_p

            st.info(f"💡 예상 소요 원료중량: 약 **{calc_tot_kg:,.2f} kg** (장당: {calc_single_g:.3f}g) | 공급가액: **{total_price:,.0f} 원**")

            submit_order = st.form_submit_button("신규 수주서 확정 등록")
            if submit_order:
                order_code = f"ORD-{datetime.datetime.now().strftime('%Y%m%d')}-{np.random.randint(100, 999)}"
                execute_query("""
                    INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    order_code, client, mat, thickness, width, length, fold, printing,
                    qty, unit_p, total_price, str(datetime.date.today()), str(due), "수주접수"
                ))
                st.success(f"수주 등록 완료! 관리번호: {order_code}")
                st.rerun()

    with tab_list:
        df_all_orders = fetch_data("SELECT * FROM orders ORDER BY order_date DESC")
        st.write(f"총 {len(df_all_orders)}건의 수주 내역이 조회되었습니다.")
        
        c_filter1, c_filter2 = st.columns([3, 7])
        with c_filter1:
            stat_choice = st.multiselect("상태 필터링", options=df_all_orders['status'].unique(), default=df_all_orders['status'].unique())
        
        filtered_orders = df_all_orders[df_all_orders['status'].isin(stat_choice)]
        st.dataframe(filtered_orders, use_container_width=True)

        st.subheader("수주 상태 업데이트")
        u_col1, u_col2, u_col3 = st.columns([3, 3, 2])
        with u_col1:
            sel_order = st.selectbox("변경할 주문코드 선택", df_all_orders['order_id'].tolist())
        with u_col2:
            new_status = st.selectbox("변경 상태값", ["수주접수", "압출대기", "생산중", "가공완료", "출고대기", "납품완료", "주문취소"])
        with u_col3:
            st.write(" ")
            if st.button("상태 갱신 적용"):
                execute_query("UPDATE orders SET status = ? WHERE order_id = ?", (new_status, sel_order))
                st.success(f"{sel_order} 주문 상태가 [{new_status}](으)로 업데이트 되었습니다.")
                st.rerun()

elif menu == "⚙️ 공정 및 생산 관리":
    st.subheader("⚙️ 비닐 공정 트래킹 (압출 -> 인쇄 -> 가공)")
    
    prod_tab1, prod_tab2 = st.tabs(["작업지시 현황", "공정 지시서 발행"])
    
    with prod_tab1:
        prod_data = fetch_data("""
            SELECT p.plan_id, p.order_id, o.client_name, o.material, p.line_no, 
                   p.process_step, p.target_qty, p.completed_qty, p.worker, p.status 
            FROM production_schedule p
            JOIN orders o ON p.order_id = o.order_id
        """)
        
        if not prod_data.empty:
            prod_data['달성률(%)'] = (prod_data['completed_qty'] / prod_data['target_qty'] * 100).round(1)
            st.dataframe(prod_data, use_container_width=True)

            st.write("#### 🛠️ 실시간 생산 실적 기입")
            w_col1, w_col2, w_col3 = st.columns(3)
            with w_col1:
                t_plan = st.selectbox("지시 번호", prod_data['plan_id'].tolist())
            with w_col2:
                add_qty = st.number_input("추가 생산 실적 수량", min_value=0, max_value=500000, step=500)
            with w_col3:
                update_step_status = st.selectbox("공정 진행상태", ["대기", "진행중", "완료", "설비점검"])

            if st.button("실적 및 상태 저장"):
                execute_query("""
                    UPDATE production_schedule 
                    SET completed_qty = completed_qty + ?, status = ?
                    WHERE plan_id = ?
                """, (add_qty, update_step_status, t_plan))
                st.success("실적이 반영되었습니다.")
                st.rerun()
        else:
            st.info("등록된 생산 공정 지시서가 없습니다.")

    with prod_tab2:
        with st.form("create_prod_plan"):
            st.write("#### 📑 신규 작업 지시서 추가")
            open_orders = fetch_data("SELECT order_id, client_name, order_qty FROM orders WHERE status != '납품완료'")
            
            p_c1, p_c2 = st.columns(2)
            with p_c1:
                if not open_orders.empty:
                    order_ref = st.selectbox("연동 수주번호", open_orders['order_id'].tolist())
                else:
                    order_ref = "연동 가능한 수주 없음"
                line = st.selectbox("배정 설비 라인", [
                    "압출 1호기 (인플레이션 광폭)", "압출 2호기 (소폭/고속)", "압출 3호기 (HD전용)",
                    "인쇄 1호기 (동판 2도)", "인쇄 2호기 (동판 4도 후렉소)",
                    "가공실링 A기 (사이드실링)", "가공실링 B기 (바텀실링/M가공)"
                ])
                step = st.selectbox("공정 단계", ["1단계 원단 압출", "2단계 그라비아 인쇄", "3단계 실링/타공 가공", "품질 전수검사"])

            with p_c2:
                target = st.number_input("지시 목표 수량", min_value=100, max_value=1000000, value=10000, step=1000)
                worker = st.selectbox("담당 작업반장", ["강기사 (압출반)", "윤반장 (인쇄반)", "송조장 (가공반)", "최기사 (검수/포장)"])
                s_date = st.date_input("작업 예정일", datetime.date.today())

            if st.form_submit_button("공정 작업지시 발행"):
                new_pid = f"PLN-{datetime.datetime.now().strftime('%m%d')}-{np.random.randint(100,999)}"
                execute_query("""
                    INSERT INTO production_schedule VALUES (?,?,?,?,?,?,?,?,?,?)
                """, (new_pid, order_ref, line, step, target, 0, worker, str(s_date), "", "진행중"))
                st.success(f"작업지시 등록 완료 ({new_pid})")
                st.rerun()

elif menu == "📦 원자재 및 완제품 재고":
    st.subheader("📦 수지 원료 및 완제품 롤 재고 현황")

    df_inv = fetch_data("SELECT * FROM inventory")

    for idx, row in df_inv.iterrows():
        if row['current_stock'] <= row['safety_stock']:
            st.error(f"⚠️ [재고경고] {row['item_name']} : 현재재고 {row['current_stock']:,}{row['unit']} (안전재고: {row['safety_stock']:,}{row['unit']})")

    st.dataframe(df_inv, use_container_width=True)

    inv_col1, inv_col2 = st.columns(2)
    with inv_col1:
        with st.expander("➕ 원자재/부자재 입고 및 출고 처리"):
            target_item = st.selectbox("대상 품목 코드", df_inv['item_code'].tolist())
            io_type = st.radio("구분", ["입고 (+)", "출고/사용 (-)"], horizontal=True)
            io_qty = st.number_input("수량 입력", min_value=1.0, max_value=100000.0, step=10.0)

            if st.button("재고 입출고 확정"):
                mod = io_qty if io_type == "입고 (+)" else -io_qty
                execute_query("""
                    UPDATE inventory 
                    SET current_stock = current_stock + ?, last_updated = ?
                    WHERE item_code = ?
                """, (mod, str(datetime.date.today()), target_item))
                st.success(f"{target_item} 재고 수량이 변경되었습니다.")
                st.rerun()

    with inv_col2:
        with st.expander("🆕 새 원자재/부자재 품목 등록"):
            with st.form("new_item_form"):
                n_code = st.text_input("품목코드 (예: RAW-PP-02)")
                n_name = st.text_input("품목명 (예: 투명 PP 원료)")
                n_cat = st.selectbox("카테고리", ["원자재", "부자재", "완제품", "반제품"])
                n_spec = st.text_input("규격 및 메이커")
                n_init = st.number_input("초기 재고", min_value=0.0, step=100.0)
                n_safe = st.number_input("기준 안전재고", min_value=0.0, step=100.0)
                n_unit = st.selectbox("단위", ["kg", "롤", "매", "BOX", "통"])
                n_cost = st.number_input("단가 (원)", min_value=0.0, step=100.0)

                if st.form_submit_button("신규 마스터 등록"):
                    if n_code and n_name:
                        execute_query("""
                            INSERT INTO inventory VALUES (?,?,?,?,?,?,?,?,?)
                        """, (n_code, n_name, n_cat, n_spec, n_init, n_safe, n_unit, n_cost, str(datetime.date.today())))
                        st.success("새 품목이 등록되었습니다.")
                        st.rerun()
                    else:
                        st.error("품목코드와 품목명은 필수입니다.")

elif menu == "👥 거래처 및 출하 관리":
    st.subheader("👥 고객사 정보 및 납품 출하 관리")

    c_tab1, c_tab2 = st.tabs(["거래처 목록", "출하증명 명세서 시뮬레이션"])

    with c_tab1:
        df_clients = fetch_data("SELECT * FROM clients")
        st.dataframe(df_clients, use_container_width=True)

        with st.expander("새 거래처 등록"):
            with st.form("add_client_form"):
                cl_code = f"CLI-{np.random.randint(100, 999)}"
                cl_name = st.text_input("고객사 / 상호명")
                cl_contact = st.text_input("담당자 성함")
                cl_phone = st.text_input("연락처")
                cl_bno = st.text_input("사업자 등록번호")
                cl_addr = st.text_input("배송지 및 공장 주소")

                if st.form_submit_button("거래처 등록 저장"):
                    if cl_name:
                        execute_query("INSERT INTO clients VALUES (?,?,?,?,?,?)",
                                      (cl_code, cl_name, cl_contact, cl_phone, cl_bno, cl_addr))
                        st.success("신규 고객사가 등록되었습니다.")
                        st.rerun()

    with c_tab2:
        ready_orders = fetch_data("SELECT * FROM orders WHERE status IN ('가공완료', '출고대기', '납품완료')")
        if not ready_orders.empty:
            sel_invoice_order = st.selectbox("출하증명서 생성 주문 건 선택", ready_orders['order_id'].tolist())
            curr_ord = ready_orders[ready_orders['order_id'] == sel_invoice_order].iloc[0]

            st.write("---")
            st.markdown(f"""
            ### 🚚 요정비닐 출하 및 납품 확인서
            - **발주일자**: {curr_ord['order_date']} | **출하일자**: {datetime.date.today()}
            - **수하인(거래처)**: {curr_ord['client_name']}
            - **제품 사양**: {curr_ord['material']} 비닐 ({curr_ord['thickness_t']}T × {curr_ord['width_w']}W × {curr_ord['length_l']}L)
            - **M가공/인쇄**: M접이 {curr_ord['fold_m']}mm / {curr_ord['printing']}
            - **출하 수량**: **{curr_ord['order_qty']:,}** 매
            - **공급 총액**: **{curr_ord['total_amount']:,.0f}** 원 (부가세 별도)
            - **인수자 서명란**: __________________ (인)
            """)
            st.write("---")
            st.button("🖨️ 거래명세표 출력 포맷 전송")
        else:
            st.info("현재 출하 가능한 가공 완료 상태의 주문이 없습니다.")

elif menu == "🧮 비닐 중량/원가 계산기":
    st.subheader("🧮 비닐 원단 규격별 중량 및 생산 견적 산출기")
    st.write("필름 사양을 입력하면 1장당 중량, 1,000장 기준 중량 및 원자재 예상 단가를 산출합니다.")

    c1, c2, c3 = st.columns(3)
    with c1:
        calc_mat = st.selectbox("수지 비중 선택", [
            ("LDPE (저밀도 폴리에틸렌 - 0.92)", 0.92),
            ("LLDPE (선형 저밀도 - 0.925)", 0.925),
            ("HDPE (고밀도 폴리에틸렌 - 0.95)", 0.95),
            ("PP (폴리프로필렌 - 0.91)", 0.91),
            ("OPP (연신 폴리프로필렌 - 0.90)", 0.90)
        ], format_func=lambda x: x[0])
        density = calc_mat[1]
        calc_t = st.number_input("두께 (T mm)", value=0.05, step=0.005, format="%.3f")

    with c2:
        calc_w = st.number_input("폭 (가로 W mm)", value=300.0, step=10.0)
        calc_l = st.number_input("길이 (세로 L mm)", value=450.0, step=10.0)
        calc_m = st.number_input("M접이 여부 (한쪽 폭 mm)", value=0.0, step=5.0)

    with c3:
        resin_cost_per_kg = st.number_input("원료 펠릿 단가 (원/kg)", value=1850, step=50)
        loss_rate = st.slider("공정 로스율 (%)", min_value=1.0, max_value=20.0, value=5.0, step=0.5)

    effective_w = calc_w + (calc_m * 2)
    single_g, _ = calculate_vinyl_weight(calc_t, effective_w, calc_l, 1, density)
    weight_per_1000 = single_g * 1000 / 1000.0

    raw_cost_per_unit = (single_g / 1000.0) * resin_cost_per_kg * (1 + (loss_rate / 100.0))

    st.divider()
    res1, res2, res3, res4 = st.columns(4)
    res1.metric("장당 순중량", f"{single_g:.3f} g")
    res2.metric("1,000장당 원단중량", f"{weight_per_1000:.2f} kg")
    res3.metric("장당 순원료비 (로스포함)", f"{raw_cost_per_unit:.2f} 원")
    res4.metric("추천 기본 출고가 (마진35%)", f"{(raw_cost_per_unit * 1.35):.2f} 원")

    st.caption("※ 계산 공식: 두께(mm) × 실질폭(cm) × 길이(cm) × 비중 × 2 (양면 백 기준) ÷ 100")