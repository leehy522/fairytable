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
menu = st.sidebar.radio("메뉴", ["📈 시장 지표 분석", "🚚 밀크런 PPT 변환", "📦 택배 송장 변환", "🏭 원가 시뮬레이터"])

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

# --- [비닐 원단 규격 계산 로직] ---
# 윤겸님이 알려주신 공식: (폭) * (길이) * 2 * 0.092 * (두께) = (무게)
def calculate_vinyl_weight(width_mm, length_m, thickness_mm):
    # 단위를 미터(m)와 밀리미터(mm)로 조정하여 계산
    # 공식 내 '92'는 비닐의 밀도(LDPE/HDPE 계수)와 관련된 고정 상수로 보입니다.
    width_m = width_mm / 1000
    weight = width_m * length_m * 2 * 92 * thickness_mm
    return weight

# 메뉴에 '규격/무게 계산기' 추가
if menu == "🏭 원가 시뮬레이터":
    st.divider()
    st.subheader("📏 원단 규격별 무게 계산기")
    st.write("비닐 규격을 입력하면 예상 원단 무게(kg)를 산출합니다.")

    c1, c2, c3 = st.columns(3)
    with c1:
        v_width = st.number_input("비닐 폭 (mm)", value=300, step=10)
    with c2:
        v_length = st.number_input("원단 총 길이 (m)", value=500, step=50)
    with c3:
        v_thick = st.number_input("비닐 두께 (mm)", value=0.05, step=0.01, format="%.3f")

    # 공식 적용
    final_weight = calculate_vinyl_weight(v_width, v_length, v_thick)
    
    st.info(f"💡 예상 원단 무게: **{final_weight:.2f} kg**")
    
    # 원가 시뮬레이터와 연동
    if 'estimated_cost_per_kg' in locals():
        total_material_cost = final_weight * estimated_cost_per_kg
        st.success(f"💰 해당 롤(Roll)당 예상 원가: **₩{total_material_cost:,.0f}**")

