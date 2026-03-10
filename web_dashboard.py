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

    # [원재료 혼합 단가]
    st.divider()
    st.subheader("🧪 원재료 혼합 단가 계산")
    # (기존의 단가 계산 로직 실행)
    st.metric("최종 제조 원가", "데이터를 입력하세요")

# --- 메뉴 5: 시장 지표 분석 ---
elif menu == "📈 시장 지표 분석":
    st.title("📈 시장 지표 정밀 분석")

    # --- [STEP 1: 상단 실시간 지표 표시] ---
    if "market_data" in st.session_state and st.session_state.market_data is not None:
        df_now = st.session_state.market_data
        try:
            # 마지막 유효 수치 추출 (NaN 제외)
            last_wti = df_now["WTI 유가"].dropna().iloc[-1]
            last_fx = df_now["원/달러 환율"].dropna().iloc[-1]
            
            # 전일 대비 변동폭 계산
            wti_delta = last_wti - df_now["WTI 유가"].dropna().iloc[-2]
            fx_delta = last_fx - df_now["원/달러 환율"].dropna().iloc[-2]

            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.metric("현재 WTI 국제유가", f"${last_wti:.2f}", f"{wti_delta:+.2f}")
            with col_m2:
                st.metric("현재 원/달러 환율", f"₩{last_fx:,.2f}", f"{fx_delta:+.2f}")
            st.divider()
        except:
            pass

    # --- [STEP 2: 데이터 동기화 (개별 호출 방식)] ---
    if st.button("🔄 최신 데이터 불러오기 (24년 포함 2년치)"):
        with st.spinner('금융 서버에서 데이터를 안전하게 가져오는 중...'):
            try:
                # 💡 리스트 형태(['CL=F', 'KRW=X']) 대신 개별 다운로드로 충돌 방지
                # 25년 3월 이전 데이터 조회를 위해 기간을 2y로 설정
                wti_raw = yf.download("CL=F", period="2y", interval="1d")['Close']
                fx_raw = yf.download("KRW=X", period="2y", interval="1d")['Close']
                
                # 데이터 병합 및 결측치 보정
                combined_df = pd.DataFrame({"WTI 유가": wti_raw, "원/달러 환율": fx_raw})
                combined_df = combined_df.ffill().bfill()
                
                if not combined_df.empty:
                    st.session_state.market_data = combined_df
                    st.success("✅ 지표 연동 성공! 이제 기간을 설정하세요.")
                    st.rerun() # 화면 즉시 갱신
                else:
                    st.error("데이터를 가져왔으나 내용이 비어있습니다.")
            except Exception as e:
                st.error(f"⚠️ 데이터 연동 실패: {e}\n잠시 후 다시 시도해 주세요.")

    # --- [STEP 3: 기간 설정 및 분석 결과] ---
    if "market_data" in st.session_state and st.session_state.market_data is not None:
        df = st.session_state.market_data
        
        col1, col2 = st.columns(2)
        with col1:
            # 24년 데이터부터 선택 가능
            start_date = st.date_input("📅 조회 시작일", df.index.max().date() - pd.Timedelta(days=14))
        with col2:
            end_date = st.date_input("📅 조회 종료일", df.index.max().date())

        mask = (df.index.date >= start_date) & (df.index.date <= end_date)
        filtered_df = df.loc[mask]

        if not filtered_df.empty:
            # 1. 시각화 그래프
            fig, ax1 = plt.subplots(figsize=(10, 4))
            ax2 = ax1.twinx()
            ax1.plot(filtered_df.index, filtered_df["WTI 유가"], color='tab:blue', label='WTI', linewidth=2)
            ax2.plot(filtered_df.index, filtered_df["원/달러 환율"], color='tab:red', label='FX', linestyle='--', linewidth=2)
            ax1.set_ylabel("WTI Price (USD)", color='tab:blue')
            ax2.set_ylabel("Exchange Rate (KRW)", color='tab:red')
            plt.title("Market Trends Analysis")
            st.pyplot(fig)

            # 2. 접이식 세부 데이터 표
            with st.expander("📝 상세 데이터 수치 보기", expanded=False):
                st.dataframe(
                    filtered_df.sort_index(ascending=False).style.format({
                        "WTI 유가": "${:.2f}", 
                        "원/달러 환율": "₩{:,.2f}"
                    }),
                    use_container_width=True
                )
        else:
            st.warning("선택하신 기간에 데이터가 없습니다.")
