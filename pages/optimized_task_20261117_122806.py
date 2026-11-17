import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

st.set_page_config(
    page_title="요정비닐 ERP - 통합 관제 대시보드",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .stMetric {
        background-color: #ffffff;
        padding: 18px 22px;
        border-radius: 12px;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.05);
        border: 1px solid #eef0f3;
        transition: transform 0.15s ease-in-out;
    }
    .stMetric:hover {
        transform: translateY(-2px);
    }
    
    .status-badge-running {
        background-color: #e8f5e9;
        color: #2e7d32;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }
    .status-badge-warning {
        background-color: #fff3e0;
        color: #e65100;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }
    .status-badge-danger {
        background-color: #ffebee;
        color: #c62828;
        padding: 4px 10px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.82rem;
        display: inline-block;
    }

    .card-container {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.04);
        border: 1px solid #eef0f3;
        margin-bottom: 20px;
    }
    
    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=600)
def generate_master_data():
    np.random.seed(42)
    end_date = datetime.today()
    start_date = end_date - timedelta(days=29)
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    
    products = [
        "HDPE 롤백(식품용)", "LDPE 지퍼백(산업용)", "OPP 접착봉투",
        "생분해성 PLA 봉투", "수축필름 25um", "알루미늄 3중 파우치"
    ]
    lines = ["압출 1호기", "압출 2호기", "인쇄 1호기", "인쇄 2호기", "가공/제단 1호기", "가공/제단 2호기"]
    clients = ["(주)신선유통", "한강패키징", "로켓물류상사", "그린푸드시스템", "글로벌팩테크", "영남식자재"]

    production_records = []
    for d in date_range:
        for line in lines:
            target_kg = np.random.uniform(1400, 2200)
            actual_kg = target_kg * np.random.uniform(0.88, 1.03)
            defect_kg = actual_kg * np.random.uniform(0.008, 0.038)
            power_kwh = actual_kg * np.random.uniform(0.42, 0.48)
            production_records.append({
                "일자": d,
                "설비": line,
                "목표량_kg": round(target_kg, 1),
                "생산량_kg": round(actual_kg, 1),
                "불량량_kg": round(defect_kg, 1),
                "전력소비_kwh": round(power_kwh, 1),
                "가동시간_h": round(np.random.uniform(18.5, 23.5), 1)
            })
    df_production = pd.DataFrame(production_records)

    sales_records = []
    for d in date_range:
        daily_order_count = np.random.randint(4, 12)
        for _ in range(daily_order_count):
            qty_roll = np.random.randint(20, 300)
            unit_price = np.random.choice([12000, 18500, 24000, 31000, 45000])
            sales_records.append({
                "수주일자": d,
                "고객사": np.random.choice(clients),
                "품목": np.random.choice(products),
                "수주수량_롤": qty_roll,
                "수주금액": qty_roll * unit_price,
                "납기상태": np.random.choice(["납기준수", "생산중", "출고대기", "지연주의"], p=[0.75, 0.15, 0.07, 0.03])
            })
    df_sales = pd.DataFrame(sales_records)

    raw_materials = [
        {"원료명": "HDPE 수지(Resin) 6000", "분류": "원자재(펠릿)", "현재고_kg": 18400, "안전재고_kg": 12000, "단가_kg": 1450},
        {"원료명": "LDPE 수지 2100", "분류": "원자재(펠릿)", "현재고_kg": 9200, "안전재고_kg": 10000, "단가_kg": 1620},
        {"원료명": "생분해 컴파운드(PLA)", "분류": "원자재(펠릿)", "현재고_kg": 4300, "안전재고_kg": 5000, "단가_kg": 2850},
        {"원료명": "백색 마스터배치(M/B)", "분류": "착색제/부자재", "현재고_kg": 2400, "안전재고_kg": 1500, "단가_kg": 3200},
        {"원료명": "그라비아 친환경 수성잉크(블랙)", "분류": "부자재", "현재고_kg": 650, "안전재고_kg": 500, "단가_kg": 7800},
        {"원료명": "지퍼테이프 12mm", "분류": "가공부자재", "현재고_kg": 28000, "안전재고_kg": 15000, "단가_m": 85}
    ]
    df_raw = pd.DataFrame(raw_materials)

    return df_production, df_sales, df_raw

df_prod, df_sales, df_raw = generate_master_data()

with st.sidebar:
    st.markdown("## 🏭 **요정비닐 ERP**")
    st.caption("스마트 팩토리 통합 관제 센터 v2.4")
    st.markdown("---")

    selected_factory = st.selectbox("🏭 제조 사업장", ["제1 화성 압출/인쇄 공장", "제2 평택 제단/가공 센터"])
    
    date_filter = st.date_input(
        "📅 조회 기간 설정",
        value=(datetime.today() - timedelta(days=14), datetime.today()),
        max_value=datetime.today()
    )
    
    if isinstance(date_filter, tuple) and len(date_filter) == 2:
        start_d, end_d = date_filter
    else:
        start_d = datetime.today() - timedelta(days=14)
        end_d = datetime.today()

    st.markdown("---")
    st.markdown("### 🔔 현장 위험 알림")
    st.markdown("""
    - <span class='status-badge-danger'>압출 2호기</span> 실린더 온도 편차 (+4.2°C)
    - <span class='status-badge-warning'>LDPE 원료</span> 안전재고 미달 위험 (-8%)
    - <span class='status-badge-running'>인쇄 1호기</span> 정기 롤러 세척 완료
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    refresh_btn = st.button("🔄 실시간 데이터 동기화", use_container_width=True)

df_prod_filtered = df_prod[(df_prod["일자"] >= pd.to_datetime(start_d)) & (df_prod["일자"] <= pd.to_datetime(end_d))]
df_sales_filtered = df_sales[(df_sales["수주일자"] >= pd.to_datetime(start_d)) & (df_sales["수주일자"] <= pd.to_datetime(end_d))]

total_sales_val = df_sales_filtered["수주금액"].sum()
total_prod_kg = df_prod_filtered["생산량_kg"].sum()
total_defect_kg = df_prod_filtered["불량량_kg"].sum()
defect_rate = (total_defect_kg / total_prod_kg * 100) if total_prod_kg > 0 else 0
total_kwh = df_prod_filtered["전력소비_kwh"].sum()
kwh_per_ton = (total_kwh / (total_prod_kg / 1000)) if total_prod_kg > 0 else 0

st.title("요정비닐 MES / ERP 통합 경영 관제 대시보드")
st.caption(f"조회 데이터 기준: {start_d} ~ {end_d} | 선택 공장: {selected_factory}")

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("총 수주 매출액", f"{total_sales_val / 1e8:.2f} 억원", delta="+6.4% MoM")
col2.metric("총 압출·가공 생산량", f"{total_prod_kg / 1e3:,.1f} 톤", delta="+2.8% 목표초과")
col3.metric("평균 종합 불량률", f"{defect_rate:.2f} %", delta="-0.34%p 개선", delta_color="inverse")
col4.metric("원단 톤당 전력 소비", f"{kwh_per_ton:,.0f} kWh/t", delta="-15 kWh 절감", delta_color="inverse")
raw_risk_count = len(df_raw[df_raw["현재고_kg"] < df_raw["안전재고_kg"]])
col5.metric("원자재 발주 필요 품목", f"{raw_risk_count} 건", delta="발주 점검 요망" if raw_risk_count > 0 else "안정", delta_color="inverse")

tab1, tab2, tab3, tab4 = st.tabs([
    "📈 경영 & 수주 영업 분석",
    "⚙️ 생산 라인 및 설비 모니터링",
    "📦 원자재 및 재고 관리 (Resin/Film)",
    "🎯 품질 / 불량 심층 분석"
])

with tab1:
    st.markdown("<div class='section-title'>📊 수주 추이 및 제품군별 포트폴리오 기여도</div>", unsafe_allow_html=True)
    c1, c2 = st.columns([7, 5])
    
    with c1:
        daily_sales = df_sales_filtered.groupby("수주일자")["수주금액"].sum().reset_index()
        fig_sales = px.area(
            daily_sales, x="수주일자", y="수주금액",
            title="일자별 수주금액 변동 추이",
            labels={"수주일자": "일자", "수주금액": "수주 총액(원)"},
            color_discrete_sequence=["#1E88E5"]
        )
        fig_sales.update_layout(hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20), height=340)
        st.plotly_chart(fig_sales, use_container_width=True)
        
    with c2:
        prod_sales = df_sales_filtered.groupby("품목")["수주금액"].sum().reset_index()
        fig_donut = px.pie(
            prod_sales, values="수주금액", names="품목", hole=0.52,
            title="비닐 제품군별 수주 비중",
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_donut.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=340)
        st.plotly_chart(fig_donut, use_container_width=True)

    c3, c4 = st.columns([6, 6])
    with c3:
        client_summary = df_sales_filtered.groupby("고객사").agg(
            수주총액=("수주금액", "sum"),
            주문건수=("품목", "count"),
            총수량_롤=("수주수량_롤", "sum")
        ).reset_index().sort_values(by="수주총액", ascending=False)
        
        fig_client = px.bar(
            client_summary, x="수주총액", y="고객사", orientation='h',
            title="주요 고객사별 매출 순위",
            labels={"수주총액": "누적 수주액(원)", "고객사": "거래처"},
            color="수주총액",
            color_continuous_scale="Blues"
        )
        fig_client.update_layout(margin=dict(l=20, r=20, t=40, b=20), height=320, yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_client, use_container_width=True)

    with c4:
        st.markdown("**📋 실시간 주요 수주 및 납기 모니터링**")
        status_disp = df_sales_filtered.sort_values(by="수주일자", ascending=False).head(7)[
            ["수주일자", "고객사", "품목", "수주수량_롤", "납기상태"]
        ]
        status_disp["수주일자"] = status_disp["수주일자"].dt.strftime('%Y-%m-%d')
        st.dataframe(status_disp, use_container_width=True, hide_index=True)

with tab2:
    st.markdown("<div class='section-title'>⚙️ 압출기 및 가공 라인 실시간 가동 현황</div>", unsafe_allow_html=True)
    
    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    line_summary = df_prod_filtered.groupby("설비").agg(
        총생산=("생산량_kg", "sum"),
        목표생산=("목표량_kg", "sum"),
        총불량=("불량량_kg", "sum"),
        가동시간=("가동시간_h", "mean")
    ).reset_index()
    line_summary["달성률"] = (line_summary["총생산"] / line_summary["목표생산"]) * 100
    line_summary["불량률"] = (line_summary["총불량"] / line_summary["총생산"]) * 100

    with col_kpi1:
        avg_oee = line_summary["달성률"].mean()
        fig_gauge = go.Figure(go.Indicator(
            mode = "gauge+number",
            value = avg_oee,
            title = {'text': "공장 전체 평균 생산 달성률"},
            gauge = {
                'axis': {'range': [70, 110]},
                'bar': {'color': "#2563EB"},
                'steps': [
                    {'range': [70, 85], 'color': "#fee2e2"},
                    {'range': [85, 95], 'color': "#fef3c7"},
                    {'range': [95, 110], 'color': "#dcfce7"}
                ],
                'threshold': {'line': {'color': "black", 'width': 3}, 'thickness': 0.8, 'value': 100}
            }
        ))
        fig_gauge.update_layout(height=260, margin=dict(l=30, r=30, t=40, b=20))
        st.plotly_chart(fig_gauge, use_container_width=True)

    with col_kpi2:
        st.markdown("**설비별 가동 지표 요약**")
        st.dataframe(
            line_summary[["설비", "총생산", "달성률", "불량률"]].style.format({
                "총생산": "{:,.0f} kg",
                "달성률": "{:.1f}%",
                "불량률": "{:.2f}%"
            }),
            use_container_width=True,
            height=220
        )

    with col_kpi3:
        fig_energy = px.scatter(
            df_prod_filtered, x="생산량_kg", y="전력소비_kwh", color="설비",
            title="생산량 대비 전력 소비 상관관계",
            opacity=0.7
        )
        fig_energy.update_layout(height=260, margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_energy, use_container_width=True)

    st.markdown("**라인별 일자별 압출 생산 추이 (kg)**")
    daily_line_prod = df_prod_filtered.groupby(["일자", "설비"])["생산량_kg"].sum().reset_index()
    fig_line_trend = px.line(
        daily_line_prod, x="일자", y="생산량_kg", color="설비",
        markers=True,
        title="라인별 일별 생산 캐파 변동 추이"
    )
    fig_line_trend.update_layout(height=340, hovermode="x unified", margin=dict(l=20, r=20, t=40, b=20))
    st.plotly_chart(fig_line_trend, use_container_width=True)

with tab3:
    st.markdown("<div class='section-title'>📦 수지(Resin) 펠릿 및 원부자재 재고 안전도 분석</div>", unsafe_allow_html=True)
    
    col_mat1, col_mat2 = st.columns([7, 5])
    
    with col_mat1:
        df_raw_chart = df_raw.copy()
        df_raw_chart["충족률(%)"] = (df_raw_chart["현재고_kg"] / df_raw_chart["안전재고_kg"]) * 100
        
        fig_stock = go.Figure()
        fig_stock.add_trace(go.Bar(
            name='현재고 (kg/m)',
            x=df_raw_chart['원료명'],
            y=df_raw_chart['현재고_kg'],
            marker_color='#3b82f6'
        ))
        fig_stock.add_trace(go.Bar(
            name='안전재고 기준 (kg/m)',
            x=df_raw_chart['원료명'],
            y=df_raw_chart['안전재고_kg'],
            marker_color='#cbd5e1'
        ))
        fig_stock.update_layout(
            barmode='group',
            title="원부자재 현재고 대비 안전재고 수준",
            height=350,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_stock, use_container_width=True)

    with col_mat2:
        st.markdown("**🚨 원자재 발주 추천 목록**")
        df_raw_action = df_raw.copy()
        df_raw_action["부족수량"] = df_raw_action["안전재고_kg"] - df_raw_action["현재고_kg"]
        df_raw_action["상태"] = df_raw_action["부족수량"].apply(
            lambda x: "⚠️ 긴급 발주 필요" if x > 0 else "정상 (재고 안전)"
        )
        
        for _, row in df_raw_action.iterrows():
            if row["부족수량"] > 0:
                st.error(f"**{row['원료명']}**: {row['부족수량']:,} kg 부족! (현: {row['현재고_kg']:,}kg / 최소: {row['안전재고_kg']:,}kg)")
            else:
                st.success(f"**{row['원료명']}**: 정상 보유 (여유분: {abs(row['부족수량']):,} kg)")

    st.markdown("**통합 원자재 자산 가치 평가**")
    df_raw["총재고액(원)"] = df_raw.apply(
        lambda r: r["현재고_kg"] * (r["단가_kg"] if "단가_kg" in r and not pd.isna(r["단가_kg"]) else r["단가_m"]),
        axis=1
    )
    total_inventory_cost = df_raw["총재고액(원)"].sum()
    st.info(f"💡 현재 보유 중인 총 원/부자재 평가액: **{total_inventory_cost:,.0f} 원** (원자재 수급 변동 반영)")
    st.dataframe(df_raw[["원료명", "분류", "현재고_kg", "안전재고_kg", "총재고액(원)"]], use_container_width=True, hide_index=True)

with tab4:
    st.markdown("<div class='section-title'>🎯 품질 보증 및 압출·인쇄 불량 원인 분석</div>", unsafe_allow_html=True)
    
    col_q1, col_q2 = st.columns([6, 6])
    
    with col_q1:
        defect_reasons = {
            "두께 편차 (게이지 불량)": 34.2,
            "인쇄 핀트/번짐 불량": 26.5,
            "실링 접착 불량": 18.3,
            "원단 겔/이물질 혼입": 12.1,
            "코로나 처리 불량": 8.9
        }
        df_defect = pd.DataFrame(list(defect_reasons.items()), columns=["불량유형", "비율"])
        df_defect = df_defect.sort_values(by="비율", ascending=False)
        df_defect["누적비율"] = df_defect["비율"].cumsum()

        fig_pareto = go.Figure()
        fig_pareto.add_trace(go.Bar(
            x=df_defect["불량유형"], y=df_defect["비율"],
            name="불량 비중 (%)", marker_color="#ef4444"
        ))
        fig_pareto.add_trace(go.Scatter(
            x=df_defect["불량유형"], y=df_defect["누적비율"],
            name="누적 비중 (%)", yaxis="y2",
            line=dict(color="#1e3a8a", width=2.5)
        ))
        fig_pareto.update_layout(
            title="불량 원인 파레토(Pareto) 분석",
            yaxis=dict(title="불량 비중 (%)"),
            yaxis2=dict(title="누적 백분율 (%)", overlaying="y", side="right", range=[0, 105]),
            height=340,
            margin=dict(l=20, r=20, t=40, b=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_pareto, use_container_width=True)

    with col_q2:
        fig_box = px.box(
            df_prod_filtered, x="설비", y="불량량_kg",
            title="설비별 일간 불량량 분포 (이상치 모니터링)",
            color="설비"
        )
        fig_box.update_layout(height=340, margin=dict(l=20, r=20, t=40, b=20), showlegend=False)
        st.plotly_chart(fig_box, use_container_width=True)

    st.markdown("**품질 개선 종합 조치 계획 (QA/QC 로그)**")
    qa_col1, qa_col2, qa_col3 = st.columns(3)
    qa_col1.warning("⚠️ **압출 1, 2호기 에어링(Air-Ring) 점검**: 두께 편차 개선을 위한 링 노즐 청소 권고")
    qa_col2.info("ℹ️ **인쇄 잉크 점도 자동관리**: 솔벤트 휘발에 따른 점도 상승 실시간 센서 추가 검토")
    qa_col3.success("✅ **작업자 실링 열선 교체 완료**: 가공 1호기 온도 컨트롤러 캘리브레이션 통과")
    
st.markdown("---")
footer_col1, footer_col2 = st.columns([8, 4])
footer_col1.caption("© 2025 (주)요정비닐 ERP 스마트제조혁신본부 | 화성 제1공장 · 평택 제2공장 실시간 통합망")
footer_col2.caption(f"마지막 데이터 갱신 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")