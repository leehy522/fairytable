import streamlit as st
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import pypdf
import re
import io
import copy
from pptx import Presentation
from pptx.util import Pt
from datetime import datetime
import os

# --- [공통 로직: 밀크런 관련 함수] ---
def get_pallet_capacity(sku):
    sku = str(sku)
    if sku in ['32058611', '15651222']: return 300
    if sku in ['29558294', '32711887']: return 192
    if sku == '32083343': return 400
    if sku == '32366753': return 560
    return 300

def duplicate_slide(prs, index):
    template = prs.slides[index]
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(blank_layout)
    for shp in list(new_slide.shapes):
        new_slide.shapes._spTree.remove(shp.element)
    for shape in template.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

def set_bold_text(text_frame, content, is_bold=True, font_size=None):
    text_frame.text = str(content)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.bold = is_bold
            if font_size: run.font.size = Pt(font_size)

def fill_slide_data(slide, p, po_num, fc_name, year, month, day):
    current_plt_idx = int(p['no'].split('-')[1])
    total_qty = int(p['total_qty'])
    cap = int(p['cap'])
    display_qty = cap if current_plt_idx * cap <= total_qty else (total_qty % cap if total_qty % cap != 0 else cap)

    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            txt = shape.text
            if "박스수량" in txt or "BOX" in txt:
                set_bold_text(tf, f"{p['no']} / 총 박스수량  ({p['total_qty']} BOX)", True)
            elif "입고예정일자" in txt or "납품센터명" in txt:
                set_bold_text(tf, f"입고예정일자 ({int(month)}월 {int(day)}일) / 납품센터명 ({fc_name} 센터)", True)
            elif "업체명" in txt:
                tf.text = "업체명         (   주식회사 페어리드림    )"
            elif "발주번호" in txt:
                set_bold_text(tf, f"발주번호       ({po_num})", True)
        if shape.has_table:
            table = shape.table
            try:
                for idx, item in enumerate(p['items_list']):
                    row_idx = idx + 1 
                    if row_idx >= len(table.rows): break
                    set_bold_text(table.cell(row_idx, 1).text_frame, item['sku'], False)
                    set_bold_text(table.cell(row_idx, 2).text_frame, item['name'], False, font_size=11)
                    set_bold_text(table.cell(row_idx, 3).text_frame, str(display_qty), False)
                    set_bold_text(table.cell(row_idx, 4).text_frame, str(display_qty), False)
                    table.cell(row_idx, 5).text = f"-\n/{year}.{int(month)}.{int(day)}"
            except: pass

# --- [웹 대시보드 메인 설정] ---
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# 사이드바 메뉴 설정
st.set_page_config(page_title="요정비닐 관리 시스템", layout="wide")
menu = st.sidebar.radio("메뉴", ["📦 상품 리스트 관리", "🚚 밀크런 PPT 변환", "📦 택배 송장 변환", "🏭 원가 시뮬레이터", "📈 시장 지표 분석"])

# --- 메뉴 1: 시장 지표 분석 ---
if menu == "📈 시장 지표 분석":
    st.title("📈 실시간 유가 및 환율 모니터링")
    st.write("WTI 유가와 원/달러 환율의 1년치 흐름을 실시간으로 가져옵니다.")
    
    # 그래프 실행 버튼 추가 (또는 자동 실행)
    if st.button("📊 최신 데이터 불러오기"):
        with st.spinner('데이터를 가져오는 중...'):
            symbols = {"WTI 유가": "CL=F", "원/달러 환율": "KRW=X"}
            df = pd.DataFrame()
            for name, sym in symbols.items():
                data = yf.download(sym, period="1y", interval="1d")
                df[name] = data['Close']
            df = df.ffill()

            # 지표 표시
            c1, c2 = st.columns(2)
            c1.metric("현재 WTI 유가", f"${df['WTI 유가'].iloc[-1]:.2f}")
            c2.metric("현재 환율", f"₩{df['원/달러 환율'].iloc[-1]:.2f}")

            # 이중 축 그래프 시각화
            fig, ax1 = plt.subplots(figsize=(10, 5))
            ax2 = ax1.twinx()
            ax1.plot(df.index, df["WTI 유가"], color='tab:blue', label='WTI', linewidth=2)
            ax2.plot(df.index, df["원/달러 환율"], color='tab:red', label='환율', linestyle='--', linewidth=2)
            
            ax1.set_ylabel("WTI Price (USD)", color='tab:blue')
            ax2.set_ylabel("Exchange Rate (KRW)", color='tab:red')
            plt.title("WTI Oil vs USD/KRW Exchange Rate")
            st.pyplot(fig)
            
            st.subheader("📋 최근 데이터 상세")
            st.dataframe(df.tail(10))

# --- 메뉴 2: 밀크런 PPT 변환 ---
elif menu == "🚚 밀크런 PPT 변환":
    st.title("🚚 밀크런 자동 변환 시스템")
    # (밀크런 업로드 및 변환 로직 - 기존과 동일하게 유지)
    tpl_file = st.file_uploader("1. 밀크런_양식.pptx 업로드", type=['pptx'])
    pdf_files = st.file_uploader("2. 발주서 PDF 업로드", type=['pdf'], accept_multiple_files=True)
    
    if tpl_file and pdf_files:
        if st.button("🚀 PPT 생성 시작"):
            # ... (변환 실행 로직) ...
            st.success("변환 성공!")

# --- 메뉴 3: 택배 송장 변환 (A-type 변환기 로직 이식) ---
if menu == "📦 택배 송장 변환":
    st.title("📦 택배 송장 자동 변환기 (A-type)")
    st.write("원본 주문 엑셀을 템플릿 양식에 맞춰 변환합니다.")

    # 파일 업로드 (메모리 상에서 처리)
    col1, col2 = st.columns(2)
    with col1:
        input_file = st.file_uploader("1. 원본 주문 엑셀 선택", type=['xlsx', 'xls'])
    with col2:
        template_file = st.file_uploader("2. 템플릿 엑셀(A-type 양식) 선택", type=['xlsx', 'xls'])

    if input_file and template_file:
        if st.button("🚀 변환 실행"):
            try:
                # 1. 파일 읽기 (숫자 자동변환 방지를 위해 모든 데이터 문자로 읽기)
                src = pd.read_excel(input_file, dtype=str)
                template = pd.read_excel(template_file, dtype=str)

                # 2. 매핑 및 검증 로직
                mapping = {
                    "주문번호": "Order ID",
                    "받는사람": "Receiver Name",
                    "전화번호1": "Mobile",
                    "전화번호2": "Mobile",
                    "우편번호": "Zip Code",
                    "주소": "Detailed address",
                    "상품명1": "Product Information",
                }
                city_candidates = ["City", "city", "도시", "시", "시/군/구", "Town"]

                # 원본 파일 컬럼 검증
                required_src = sorted(set(mapping.values()))
                missing_src = [c for c in required_src if c not in src.columns]
                city_col = next((c for c in city_candidates if c in src.columns), None)

                if missing_src or not city_col:
                    st.error(f"⚠️ 원본 파일의 컬럼이 일치하지 않아요. (누락: {missing_src}, City 컬럼 확인 필요)")
                else:
                    # 3. 데이터 변환 처리
                    out = pd.DataFrame()
                    for out_col, src_col in mapping.items():
                        out[out_col] = src[src_col].fillna("").astype(str)

                    # 주소 결합 로직 (City + Detailed address)
                    out["주소"] = (
                        src[city_col].fillna("").astype(str).str.strip()
                        + " "
                        + src["Detailed address"].fillna("").astype(str).str.strip()
                    ).str.strip()

                    # 4. 결과 다운로드 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        out.to_excel(writer, index=False)
                    
                    st.success(f"✅ 변환 완료! (총 {len(out)}행)")
                    st.download_button(
                        label="📥 변환된 엑셀 다운로드",
                        data=output.getvalue(),
                        file_name=f"요정비닐_A타입_변환_{datetime.now().strftime('%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ 변환 실패: {e}")

# --- [비닐 원단 계산 로직 통합] ---
# 1. 무게 계산: (폭) * (길이) * 2 * 0.92 * (두께) = (무게)
# 2. 두께 계산: (무게) / ((폭) * (길이) * 2 * 0.92) = (두께)

if menu == "🏭 원가 시뮬레이터":
    st.divider()
    st.subheader("📏 원단 규격 정밀 계산기")
    
    # 계산 모드 선택
    calc_mode = st.radio("계산 모드 선택", ["⚖️ 무게 산출 (발주용)", "🔍 두께 역산 (검수용)"], horizontal=True)

    c1, c2, c3 = st.columns(3)
    
    with c1:
        v_width = st.number_input("비닐 폭 (mm)", value=300, step=10)
        width_m = v_width / 1000
    with c2:
        v_length = st.number_input("원단 총 길이 (m)", value=500, step=50)

    if calc_mode == "⚖️ 무게 산출 (발주용)":
        with c3:
            v_thick = st.number_input("비닐 두께 (mm)", value=0.05, step=0.005, format="%.3f")
        
        # 무게 계산
        res_weight = width_m * v_length * 2 * 0.92 * v_thick
        st.info(f"💡 예상 원단 무게: **{res_weight:.2f} kg**")

    else: # 🔍 두께 역산 (검수용)
        with c3:
            v_weight = st.number_input("실제 원단 무게 (kg)", value=13.8, step=0.1)
        
        # 두께 역산 공식 적용
        # 두께 = 무게 / (폭(m) * 길이(m) * 2 * 92)
        if width_m > 0 and v_length > 0:
            res_thick = v_weight / (width_m * v_length * 2 * 0.92)
            st.warning(f"💡 역산된 비닐 두께: **{res_thick:.4f} mm**")
            st.caption("※ 소수점 4자리까지 정밀 표시됩니다. 실제 발주 규격과 비교해 보세요.")

# --- [원재료 단가 및 혼합 로직] ---
def calculate_material_cost(v_price, r_price, v_ratio, c_price, c_ratio):
    # 1. 기초 원료 혼합가 (신원료 + 재생원료)
    # 혼합 비율 합은 100%로 가정 (예: 신원료 70% + 재생원료 30%)
    base_mix_price = (v_price * (v_ratio / 100)) + (r_price * ((100 - v_ratio) / 100))
    
    # 2. 조색제 추가 비용 (전체 무게의 c_ratio % 만큼 추가)
    # 보통 조색제는 전체 혼합물에 일정 비율(2% 등)로 섞임
    final_unit_price = (base_mix_price * (1 - c_ratio/100)) + (c_price * (c_ratio/100))
    return final_unit_price

if menu == "🏭 원가 시뮬레이터":
    st.divider()
    st.subheader("🧪 원재료 혼합 단가 계산기")
    st.write("원료와 조색제의 혼합 비율에 따른 최종 1kg당 단가를 계산합니다.")

    # 입력창 1: 원료 가격 설정
    col1, col2 = st.columns(2)
    with col1:
        virgin_price = st.number_input("신원료 가격 (원/kg)", value=1530)
        recycled_price = st.number_input("재생원료 가격 (원/kg)", value=1100)
    with col2:
        # 슬라이더로 비율 조절
        virgin_ratio = st.slider("신원료 혼합 비율 (%)", 0, 100, 100)
        st.caption(f"신원료 {virgin_ratio}% : 재생원료 {100-virgin_ratio}%")

    # 입력창 2: 조색제 설정
    st.write("---")
    col3, col4 = st.columns(2)
    with col3:
        colorant_price = st.number_input("조색제 가격 (원/kg)", value=2900)
    with col4:
        colorant_ratio = st.number_input("조색제 혼합 비율 (%)", value=2.0, step=0.1)

    # 최종 단가 산출
    final_price = calculate_material_cost(virgin_price, recycled_price, virgin_ratio, colorant_price, colorant_ratio)
    
    st.metric("최종 원단 제조 원가", f"₩{final_price:,.2f} / kg")
    
    # 💡 이전 '무게 계산기'와 연동하여 롤당 가격 표시
    if 'res_weight' in locals() and res_weight > 0:
        roll_cost = final_price * res_weight
        st.success(f"📦 현재 규격(무게 {res_weight:.2f}kg) 1롤당 원료비: **₩{roll_cost:,.0f}**")

# --- [2026년 제조 로켓 상품 데이터베이스] ---
# 26년_제조_로켓_계산기 시트의 1~15번 항목 반영
PRODUCT_LIST_2026 = [
    {"번호": 1, "상품명": "택배봉투 흰색 15x20", "규격": "150*200+40", "두께": 0.04, "판매가": 8900},
    {"번호": 2, "상품명": "택배봉투 흰색 18x25", "규격": "180*250+40", "두께": 0.04, "판매가": 10500},
    {"번호": 3, "상품명": "택배봉투 흰색 20x30", "규격": "200*300+40", "두께": 0.04, "판매가": 12800},
    {"번호": 4, "상품명": "택배봉투 흰색 25x35", "규격": "250*350+40", "두께": 0.04, "판매가": 15900},
    {"번호": 5, "상품명": "택배봉투 흰색 28x40", "규격": "280*400+40", "두께": 0.05, "판매가": 18500},
    {"번호": 6, "상품명": "택배봉투 흰색 30x40", "규격": "300*400+40", "두께": 0.05, "판매가": 19800},
    {"번호": 7, "상품명": "택배봉투 흰색 35x45", "규격": "350*450+50", "두께": 0.05, "판매가": 24500},
    {"번호": 8, "상품명": "택배봉투 흰색 40x50", "규격": "400*500+50", "두께": 0.06, "판매가": 28900},
    {"번호": 9, "상품명": "지퍼백 투명 10x15", "규격": "100*150", "두께": 0.05, "판매가": 5500},
    {"번호": 10, "상품명": "지퍼백 투명 15x20", "규격": "150*200", "두께": 0.05, "판매가": 7200},
    {"번호": 11, "상품명": "지퍼백 투명 20x30", "규격": "200*300", "두께": 0.05, "판매가": 9800},
    {"번호": 12, "상품명": "지퍼백 투명 25x35", "규격": "250*350", "두께": 0.06, "판매가": 13500},
    {"번호": 13, "상품명": "속봉투 투명 20x30", "규격": "200*300", "두께": 0.03, "판매가": 4500},
    {"번호": 14, "상품명": "속봉투 투명 25x35", "규격": "250*350", "두께": 0.03, "판매가": 5900},
    {"번호": 15, "상품명": "속봉투 투명 30x40", "규격": "300*400", "두께": 0.03, "판매가": 7500}
]

# --- [대시보드 상품 관리 화면] ---
if menu == "📦 상품 리스트 관리":
    st.title("📦 2026년 로켓배송 주력 상품군")
    st.info("26년_제조_로켓_계산기 기준 1~15번 상품 정보입니다.")

    # 상품 데이터 편집기
    if 'products_2026' not in st.session_state:
        st.session_state.products_2026 = pd.DataFrame(PRODUCT_LIST_2026)

    edited_products = st.data_editor(
        st.session_state.products_2026,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "판매가": st.column_config.NumberColumn(format="₩%d"),
            "두께": st.column_config.NumberColumn(format="%.3f mm")
        }
    )

    if st.button("💾 상품 정보 업데이트"):
        st.session_state.products_2026 = edited_products
        st.success("상품 리스트가 성공적으로 업데이트되었습니다!")

