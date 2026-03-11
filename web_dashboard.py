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

elif menu == "🚚 밀크런 PPT 변환":
    st.title("🚚 밀크런 자동 변환 시스템 (v4.98 이식)")
    st.info("💡 수량이 팔레트별 실제 적재량으로 자동 계산되어 PPT에 표시됩니다.") # [cite: 11]

    tpl_file = st.file_uploader("1. 밀크런_양식.pptx 업로드", type=['pptx'])
    pdf_files = st.file_uploader("2. 발주서 PDF 업로드", type=['pdf'], accept_multiple_files=True)
    
    if tpl_file and pdf_files:
        if st.button("🚀 PPT 생성 시작"):
            with st.spinner("발주서 분석 중..."):
                try:
                    prs = Presentation(tpl_file)
                    # 기존 슬라이드 정리 (첫 번째 슬라이드 제외) [cite: 26]
                    while len(prs.slides) > 1:
                        rId = prs.slides._sle[1].rId
                        prs.part.drop_rel(rId); del prs.slides._sle[1]
                    
                    extracted_data = []
                    for pdf_file in pdf_files:
                        reader = pypdf.PdfReader(pdf_file)
                        text = ""
                        for page in reader.pages: text += page.extract_text() + "\n" # [cite: 20]
                        
                        # --- [데이터 추출 로직 이식] ---
                        po_match = re.search(r"(?:발주번호|PO|no|Info)\s*[:\s\n]*(\d{9})", text, re.I) # 
                        po_num = po_match.group(1) if po_match else re.findall(r"\b\d{9}\b", text)[0]
                        
                        fc_match = re.search(r"(?:FC명|FC\s*Name|센터명)\s*[:\s\n]*([A-Z0-9가-힣]+)", text, re.I) # 
                        fc_name = fc_match.group(1).strip() if fc_match else "알수없음"
                        
                        date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text) # [cite: 22]
                        date_raw = date_match.group(1) if date_match else "2026-03-09"
                        
                        items, processed = [], set()
                        sku_matches = list(re.finditer(r"\b(\d{8})\b", text)) # [cite: 22]
                        for i, m in enumerate(sku_matches):
                            sku = m.group(1)
                            if sku in processed: continue
                            block = text[m.end():m.end()+450] # [cite: 23]
                            name_search = re.search(r"([가-힣]{2,}[가-힣\s\d\-\(\)]+)", block)
                            name = name_search.group(1).strip() if name_search else "상품명확인"
                            nums = re.findall(r"\b\d{1,4}\b", block) # [cite: 23]
                            qty = int(nums[1]) if len(nums) >= 2 else (int(nums[0]) if len(nums) == 1 else 0)
                            if qty > 0:
                                items.append({'sku': sku, 'name': name[:40], 'qty': qty, 'cap': get_pallet_capacity(sku)}) # [cite: 24]
                                processed.add(sku)
                        
                        extracted_data.append({'po': po_num, 'fc': fc_name, 'date': date_raw, 'items': items})

                    # --- [슬라이드 생성 로직 이식] ---
                    is_first = True
                    for data in extracted_data:
                        y, m, d = data['date'].split('-')
                        for item in data['items']:
                            # 팔레트 수량 계산 (분할 규칙 적용) [cite: 12, 27]
                            tot_plt = (item['qty'] // item['cap']) + (1 if item['qty'] % item['cap'] > 0 else 0)
                            
                            for i in range(1, tot_plt + 1):
                                p_info = {
                                    'no': f"{tot_plt}-{i}", # [cite: 28]
                                    'total_qty': item['qty'], 
                                    'cap': item['cap'], 
                                    'items_list': [item]
                                }
                                # 한 팔레트당 2장씩 생성 루틴 유지 [cite: 28]
                                for _ in range(2):
                                    slide = prs.slides[0] if is_first else duplicate_slide(prs, 0)
                                    is_first = False
                                    fill_slide_data(slide, p_info, data['po'], data['fc'], y, m, d) # [cite: 28]

                    # 다운로드 준비
                    ppt_out = io.BytesIO()
                    prs.save(ppt_out)
                    ppt_out.seek(0)
                    
                    st.success("✨ 수량 분할 규칙이 적용된 PPT 생성이 완료되었습니다!") # [cite: 29]
                    st.download_button(
                        label="📥 변환된 PPT 다운로드",
                        data=ppt_out.getvalue(),
                        file_name=f"밀크런_결과_{datetime.now().strftime('%m%d_%H%M')}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except Exception as e:
                    st.error(f"❌ 오류 발생: {e}")
                    
# --- 메뉴 3: 택배 송장 변환 ---
elif menu == "📦 택배 송장 변환":
    st.title("📦 택배 송장 자동 변환기 (A-type)")
    st.write("원본 주문 엑셀을 요정비닐 템플릿 양식에 맞춰 변환합니다.")

    col1, col2 = st.columns(2)
    with col1:
        input_file = st.file_uploader("1. 원본 주문 엑셀 선택", type=['xlsx', 'xls'])
    with col2:
        template_file = st.file_uploader("2. 템플릿 엑셀(A-type) 선택", type=['xlsx', 'xls'])

    if input_file and template_file:
        if st.button("🚀 변환 실행"):
            try:
                # 1. 파일 읽기 (숫자 변형 방지를 위해 문자열로 읽기)
                src = pd.read_excel(input_file, dtype=str)
                
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
                # 도시 정보가 담긴 다양한 컬럼명 후보들
                city_candidates = ["City", "city", "도시", "시", "시/군/구"]
                city_col = next((c for c in city_candidates if c in src.columns), None)

                if not city_col:
                    st.error("⚠️ 원본 파일에서 'City' 또는 '도시' 컬럼을 찾을 수 없습니다.")
                else:
                    # 3. 데이터 변환 처리
                    out = pd.DataFrame()
                    for out_col, src_col in mapping.items():
                        if src_col in src.columns:
                            out[out_col] = src[src_col].fillna("").astype(str)

                    # 주소 결합 로직: City + Detailed address
                    out["주소"] = (
                        src[city_col].fillna("").astype(str).str.strip() + " " + 
                        src["Detailed address"].fillna("").astype(str).str.strip()
                    ).str.strip()

                    # 전화번호2는 1과 동일하게 세팅
                    out["전화번호2"] = out["전화번호1"]

                    # 4. 결과 다운로드 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        out.to_excel(writer, index=False)
                    
                    st.success(f"✅ 변환 완료! (총 {len(out)}행)")
                    st.download_button(
                        label="📥 변환된 엑셀 다운로드",
                        data=output.getvalue(),
                        file_name=f"요정비닐_송장변환_{datetime.now().strftime('%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ 변환 실패: {e}")
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

    # --- [1. 원재료 혼합 단가 계산 로직] ---
    st.subheader("🧪 1. 원재료 혼합 단가 설정")
    
    col1, col2 = st.columns(2)
    with col1:
        v_price = st.number_input("신원료 가격 (원/kg)", value=1530) #
        r_price = st.number_input("재생원료 가격 (원/kg)", value=1300) #
    with col2:
        v_ratio = st.slider("신원료 혼합 비율 (%)", 0, 100, 70) #
        st.caption(f"신원료 {v_ratio}% : 재생원료 {100-v_ratio}%")

    st.write("---")
    col3, col4 = st.columns(2)
    with col3:
        c_price = st.number_input("조색제 가격 (원/kg)", value=2700) #
    with col4:
        c_ratio = st.number_input("조색제 혼합 비율 (%)", value=2.5, step=0.1, format="%.1f") #

    # 💡 핵심 계산식: (신원료혼합가) + (조색제추가비용)
    # 1. 기초 원료 혼합 단가
    base_price = (v_price * (v_ratio / 100)) + (r_price * ((100 - v_ratio) / 100))
    # 2. 조색제 포함 최종 단가
    final_unit_price = (base_price * (1 - c_ratio/100)) + (c_price * (c_ratio/100))
    
    st.success(f"🎨 **최종 원재료 단가: ₩{final_unit_price:,.2f} / kg**")

    # --- [2. 원단 규격 및 롤당 가격 계산] ---
    st.divider()
    st.subheader("📏 2. 원단 규격 및 생산 원가")
    
    c_w1, c_w2, c_w3 = st.columns(3)
    with c_w1:
        width_mm = st.number_input("비닐 폭 (mm)", value=630) #
    with c_w2:
        length_m = st.number_input("원단 총 길이 (m)", value=1800) #
    with c_w3:
        thick_mm = st.number_input("비닐 두께 (mm)", value=0.009, step=0.001, format="%.3f") #

    # 💡 원단 무게 계산 공식: (폭m * 길이m * 2 * 비중0.92 * 두께mm)
    total_weight = (width_mm / 1000) * length_m * 2 * 0.92 * thick_mm
    total_cost = total_weight * final_unit_price

    col_res1, col_res2 = st.columns(2)
    with col_res1:
        st.metric("예상 원단 무게", f"{total_weight:.2f} kg")
    with col_res2:
        st.metric("1롤당 제조 원가", f"₩{total_cost:,.0f}")
        
# --- 메뉴 5: 시장 지표 분석 (초기 복구 버전) ---
elif menu == "📈 시장 지표 분석":
    st.title("📈 실시간 유가 및 환율 모니터링")
    st.write("WTI 유가와 원/달러 환율의 1년치 흐름을 실시간으로 가져옵니다.")
    
    # 💡 초기 방식: 버튼을 누르면 즉시 yfinance에서 데이터를 내려받습니다.
    if st.button("📊 최신 데이터 불러오기"):
        with st.spinner('데이터를 가져오는 중...'):
            try:
                # 데이터 심볼 설정 (WTI: CL=F, 환율: KRW=X)
                symbols = {"WTI 유가": "CL=F", "원/달러 환율": "KRW=X"}
                df = pd.DataFrame()
                
                for name, sym in symbols.items():
                    # 기간을 1년(1y)으로 설정하여 안정적으로 가져옵니다.
                    data = yf.download(sym, period="2y", interval="1d")
                    df[name] = data['Close']
                
                # 빈 값 채우기
                df = df.ffill()

                # 1. 주요 지표 표시 (Metric)
                c1, c2 = st.columns(2)
                c1.metric("현재 WTI 유가", f"${df['WTI 유가'].iloc[-1]:.2f}")
                c2.metric("현재 환율", f"₩{df['원/달러 환율'].iloc[-1]:.2f}")

                # 2. 이중 축 그래프 시각화 (Matplotlib)
                fig, ax1 = plt.subplots(figsize=(10, 5))
                ax2 = ax1.twinx()
                
                ax1.plot(df.index, df["WTI 유가"], color='tab:blue', label='WTI', linewidth=2)
                ax2.plot(df.index, df["원/달러 환율"], color='tab:red', label='환율', linestyle='--', linewidth=2)
                
                ax1.set_ylabel("WTI Price (USD)", color='tab:blue')
                ax2.set_ylabel("Exchange Rate (KRW)", color='tab:red')
                plt.title("WTI Oil vs USD/KRW Exchange Rate")
                
                # 범례 표시
                lines1, labels1 = ax1.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
                
                st.pyplot(fig)
                
                # 3. 상세 데이터 표 (하단)
                st.subheader("📋 최근 데이터 상세")
                st.dataframe(df.tail(10).sort_index(ascending=False), use_container_width=True)
                
            except Exception as e:
                st.error(f"데이터 연동 실패: {e}")
