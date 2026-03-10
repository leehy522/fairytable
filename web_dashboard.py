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

# --- 메뉴 5: 시장 지표 분석 (보강 버전) ---
elif menu == "📈 시장 지표 분석":
    st.title("📈 실시간 유가 및 환율 모니터링")
    
    if "market_data" not in st.session_state:
        st.session_state.market_data = None

    if st.button("📊 최신 데이터 불러오기"):
        with st.spinner('금융 데이터를 가져오는 중...'):
            try:
                # 1. 데이터 다운로드 (CL=F: 유가, KRW=X: 환율)
                # 'auto_adjust=True'를 넣어 데이터 누락을 방지합니다.
                raw_data = yf.download(["CL=F", "KRW=X"], period="1y", interval="1d", auto_adjust=True)
                
                # 2. Close(종가) 데이터만 추출
                df = raw_data['Close'].rename(columns={"CL=F": "WTI 유가", "KRW=X": "원/달러 환율"})
                
                # 3. 데이터 검증 및 결측치 채우기
                if df["원/달러 환율"].isnull().all():
                    st.error("⚠️ 환율 데이터를 가져오지 못했습니다. 잠시 후 다시 시도해주세요.")
                else:
                    df = df.ffill().bfill() # 앞뒤로 빈칸을 꽉 채웁니다.
                    st.session_state.market_data = df
                    st.success("✅ 최신 데이터를 성공적으로 업데이트했습니다.")
                    
            except Exception as e:
                st.error(f"데이터 로딩 오류: {e}")

    # 데이터 출력 로직
    if st.session_state.market_data is not None:
        df = st.session_state.market_data
        
        c1, c2 = st.columns(2)
        # 마지막 행의 데이터가 NaN인지 확인 후 안전하게 가져오기
        current_wti = df["WTI 유가"].dropna().iloc[-1]
        current_ex = df["원/달러 환율"].dropna().iloc[-1]
        
        c1.metric("현재 WTI 유가", f"${current_wti:.2f}")
        c2.metric("현재 환율", f"₩{current_ex:.2f}")

        # 그래프 시각화 (기존 코드와 동일)
        fig, ax1 = plt.subplots(figsize=(10, 5))
        ax2 = ax1.twinx()
        ax1.plot(df.index, df["WTI 유가"], color='tab:blue', label='WTI')
        ax2.plot(df.index, df["원/달러 환율"], color='tab:red', label='환율', linestyle='--')
        plt.title("WTI Oil vs USD/KRW Exchange Rate")
        st.pyplot(fig)
