import streamlit as st
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import copy
import io
import pypdf
import re
from pptx import Presentation
from pptx.util import Pt
from datetime import datetime

# --- [1. 페이지 기본 설정] ---
# 이 코드가 가장 먼저 나와야 합니다.
st.set_page_config(page_title="요정비닐 스마트 시스템", layout="wide")

# --- [2. 로그인 보안 로직] ---
def check_password():
    """아이디와 비밀번호를 확인하여 로그인을 제어합니다."""
    USER_ID = "lhy"
    USER_PW = "dlghkdud1%" #

    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    # 로그인 화면 디자인
    st.title("🔐 요정비닐 시스템 접속")
    col_l, col_r = st.columns([1, 2])
    with col_l:
        input_id = st.text_input("ID", placeholder="아이디 입력", key="login_id")
        input_pw = st.text_input("Password", type="password", placeholder="비밀번호 입력", key="login_pw")
        
        if st.button("로그인 실행"):
            if input_id == USER_ID and input_pw == USER_PW:
                st.session_state.password_correct = True
                st.rerun() # 성공 시 즉시 화면 갱신
            else:
                st.error("❌ 정보가 일치하지 않습니다.")
    return False

# 로그인 체크 실행 (통과 못 하면 여기서 멈춤)
if not check_password():
    st.stop()

# --- [3. 공통 로직 및 함수 정의] ---
# (여기서부터 기존 버전 1의 get_pallet_capacity 등 함수들이 이어집니다)
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

# 💡 텍스트 공란 문제를 해결하기 위해 로직을 꽉 채웠습니다.
def fill_slide_data(slide, p, po_num, fc_name, year, month, day):
    try:
        current_plt_idx = int(p['no'].split('-')[1])
        total_qty = int(p['total_qty'])
        cap = int(p['cap'])
        # 팔레트별 수량 계산 로직 (v4.98)
        display_qty = cap if current_plt_idx * cap <= total_qty else (total_qty % cap if total_qty % cap != 0 else cap)
    except: 
        display_qty = p['total_qty']

    # 💡 윤겸님이 말씀하신 바로 그 핵심 텍스트 치환 로직입니다!
    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            txt = shape.text
            
            # 1. 박스수량 및 팔레트 번호 (예: 12-1)
            if "박스수량" in txt or "BOX" in txt:
                set_bold_text(tf, f"{p['no']} / 총 박스수량  ({p['total_qty']} BOX)", True)
            
            # 2. 입고 날짜 및 납품 센터명
            elif "입고예정일자" in txt or "납품센터명" in txt:
                set_bold_text(tf, f"입고예정일자 ({int(month)}월 {int(day)}일) / 납품센터명 ({fc_name} 센터)", True)
            
            # 3. 고정 업체명 입력
            elif "업체명" in txt:
                tf.text = "업체명         (   주식회사 페어리드림    )"
            
            # 4. 발주번호 입력
            elif "발주번호" in txt:
                set_bold_text(tf, f"발주번호       ({po_num})", True)
        
        # 표(Table)가 있는 경우 SKU와 상품명, 수량을 채웁니다.
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
                
@st.cache_data(ttl=60)
def load_google_sheet_data():
    CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTVvCbm9KEoUrqvlXSyIyLHmstIGZuiuTMLYDBnmgnxInrfoMelDXFSWogUdHUfNALb7uC_nBAIyzif/pub?output=csv"
    try:
        df = pd.read_csv(CSV_URL)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(subset=['상품명'])
    except: return pd.DataFrame()

# --- [여기서부터 기존 메뉴 로직 시작] ---
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
    st.title("🚚 밀크런 통합 편집 시스템 (v4.98 이식)")
    st.info("💡 수량을 직접 수정하거나 여러 품목을 하나로 합쳐서(Mixed) PPT를 생성할 수 있습니다.")

    # 1. 파일 업로드
    tpl_file = st.file_uploader("1. 양식 PPT 업로드", type=['pptx'], key="ml_tpl")
    pdf_files = st.file_uploader("2. 발주서 PDF 업로드 (다중 선택 가능)", type=['pdf'], accept_multiple_files=True, key="ml_pdfs")

    if tpl_file and pdf_files:
        # 데이터 추출 단계
        if "extracted_data" not in st.session_state:
            st.session_state.extracted_data = []

# 💡 버튼 아래쪽은 모두 이만큼(스페이스 8칸) 들여쓰기가 되어야 합니다.
        if st.button("🔍 발주서 데이터 정밀 분석"):
            all_extracted = []
            for pdf_file in pdf_files:
                reader = pypdf.PdfReader(pdf_file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"
                
                # 데이터 추출 로직 시작 (v4.98 정밀 로직)
                po_num = (re.search(r"(?:발주번호|PO|no|Info)\s*[:\s\n]*(\d{9})", text, re.I) or 
                          re.search(r"(\d{9})", text) or ["000000000"])[0]
                
                fc_match = re.search(r"(?:FC명|FC\s*Name|센터명)\s*[:\s\n]*([A-Z0-9가-힣]+)", text, re.I) or \
                           re.search(r"([가-힣]+)센터", text)
                fc_name = fc_match.group(1).strip() if fc_match else "알수없음"
                
                date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
                date_raw = date_match.group(1) if date_match else "2026-03-12"
                
                sku_matches = list(re.finditer(r"\b(\d{8})\b", text))
                processed = set()
                
                for m in sku_matches:
                    sku = m.group(1)
                    if sku in processed: continue
                    cap = get_pallet_capacity(sku)
                    block = text[m.end():m.end()+450]
                    name_search = re.search(r"([가-힣]{2,}[가-힣\s\d\-\(\)]+)", block)
                    real_name = name_search.group(1).strip() if name_search else "상품명확인"
                    
                    nums = re.findall(r"\b\d{1,4}\b", block)
                    qty = int(nums[1]) if len(nums) >= 2 else (int(nums[0]) if len(nums) == 1 else 0)
                    
                    # 💡 여기가 이미지에서 빨간 줄이 가있던 핵심 부분입니다!
                    all_extracted.append({
                        "발주번호": po_num, 
                        "센터": fc_name, 
                        "SKU": sku, 
                        "상품명": real_name[:40], 
                        "확정수량": qty, 
                        "적재량": cap, 
                        "date": date_raw
                    })
                    processed.add(sku)
            
            # 모든 분석이 끝나면 세션에 저장 (이것도 버튼 안쪽!)
            st.session_state.extracted_data = all_extracted
            st.rerun()

        # 2. 통합 편집기 (v4.98 BulkQuantityEditor 기능)
        if st.session_state.extracted_data:
            st.subheader("📊 발주 데이터 통합 편집")
            df_editor = pd.DataFrame(st.session_state.extracted_data)
            
            # 직접 수량과 적재량을 수정할 수 있는 편집 테이블
            edited_df = st.data_editor(df_editor, num_rows="dynamic", use_container_width=True, key="ml_editor")

            # 3. PPT 생성 시작
if st.button("🚀 지능형 합짐 및 PPT 생성"):
                try:
                    prs = Presentation(tpl_file)
                    # 기존 슬라이드 정리 로직 생략 (v1.2와 동일)
                    
                    is_first = True
                    # 1. 💡 센터별로 그룹화하여 합짐 데이터를 생성합니다.
                    # 발주번호가 달라도 '센터'가 같으면 한 그룹으로 묶입니다.
                    for center, group in edited_df.groupby("센터"):
                        
                        # 2. 이 그룹(같은 센터)의 모든 품목을 하나의 리스트에 담습니다.
                        mixed_items = []
                        total_qty_sum = 0
                        for _, row in group.iterrows():
                            if int(row["확정수량"]) <= 0: continue
                            
                            mixed_items.append({
                                'sku': row['SKU'],
                                # 발주번호를 상품명 앞에 나란히 표기하여 검수 편의성 증대
                                'name': f"[{row['발주번호']}] {row['상품명']}", 
                                'qty': int(row['확정수량'])
                            })
                            total_qty_sum += int(row['확정수량'])
                        
                        if not mixed_items: continue

                        # 3. 묶인 품목들을 팔레트 분할 규칙에 따라 슬라이드로 생성
                        # 적재량은 해당 그룹 내 품목 중 하나를 기준으로 잡습니다.
                        cap = int(group["적재량"].iloc[0]) 
                        tot_plt = (total_qty_sum // cap) + (1 if total_qty_sum % cap > 0 else 0)
                        
                        y, m, d = group["date"].iloc[0].split('-')
                        
                        for i in range(1, tot_plt + 1):
                            p_info = {
                                'no': f"{tot_plt}-{i}", 
                                'total_qty': total_qty_sum, 
                                'cap': cap, 
                                'items_list': mixed_items # 💡 여기에 여러 발주번호 품목이 나란히 들어감
                            }
                            # 한 팔레트당 2장씩 생성 루틴
                            for _ in range(2):
                                slide = prs.slides[0] if is_first else duplicate_slide(prs, 0)
                                is_first = False
                                # 발주번호 자리에는 '혼합발주' 또는 대표번호 표기
                                fill_slide_data(slide, p_info, "혼합(Mixed)", center, y, m, d)
                    ppt_out = io.BytesIO()
                    prs.save(ppt_out)
                    st.download_button("📥 최종 PPT 다운로드", ppt_out.getvalue(), "밀크런_v4.98_결과.pptx")
                    st.success("✨ 수량 분할 규칙이 적용된 PPT가 생성되었습니다!")
                except Exception as e:
                    st.error(f"오류: {e}")
                    
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
                
elif menu == "🏭 원가 시뮬레이터":
    st.title("🏭 원가 및 규격 시뮬레이터")
    
    # --- [섹션 1: 원단 규격 정밀 계산기] ---
    st.subheader("📏 원단 규격 정밀 계산기")
    # key를 추가하여 하단 위젯과 ID 충돌을 방지합니다.
    calc_mode = st.radio("계산 모드", ["⚖️ 무게 산출", "🔍 두께 역산"], horizontal=True, key="mode_selector")
    
    c1, c2, c3 = st.columns(3)
    with c1: 
        v_width = st.number_input("비닐 폭 (mm)", value=630, key="calc_v_width") # key 추가
    with c2: 
        v_length = st.number_input("원단 총 길이 (m)", value=1800, key="calc_v_length") # key 추가
    
    res_weight = 0
    if calc_mode == "⚖️ 무게 산출":
        with c3: 
            v_thick = st.number_input("두께 (mm)", value=0.009, format="%.3f", key="calc_v_thick")
        res_weight = (v_width/1000) * v_length * 2 * 0.92 * v_thick
        st.info(f"💡 예상 무게: {res_weight:.2f} kg")
    else:
        with c3: 
            v_weight_in = st.number_input("실제 무게 (kg)", value=13.8, key="calc_v_weight_in")
        res_thick = v_weight_in / ((v_width/1000) * v_length * 2 * 0.92)
        st.warning(f"💡 역산된 두께: {res_thick:.4f} mm")

    # --- [섹션 2: 원재료 혼합 단가 계산] ---
    st.divider()
    st.subheader("🧪 1. 원재료 혼합 단가 설정")
    
    col1, col2 = st.columns(2)
    with col1:
        v_price = st.number_input("신원료 가격 (원/kg)", value=1530, key="raw_v_price") 
        r_price = st.number_input("재생원료 가격 (원/kg)", value=1300, key="raw_r_price") 
    with col2:
        v_ratio = st.slider("신원료 혼합 비율 (%)", 0, 100, 70, key="raw_v_ratio") 
        st.caption(f"신원료 {v_ratio}% : 재생원료 {100-v_ratio}%")

    st.write("---")
    col3, col4 = st.columns(2)
    with col3:
        c_price = st.number_input("조색제 가격 (원/kg)", value=2700, key="raw_c_price") 
    with col4:
        c_ratio = st.number_input("조색제 혼합 비율 (%)", value=2.5, step=0.1, format="%.1f", key="raw_c_ratio") 

    # 혼합 단가 수식 유지
    base_price = (v_price * (v_ratio / 100)) + (r_price * ((100 - v_ratio) / 100))
    final_unit_price = (base_price * (1 - c_ratio/100)) + (c_price * (c_ratio/100))
    st.success(f"🎨 **최종 원재료 단가: ₩{final_unit_price:,.2f} / kg**")

    # --- [섹션 3: 원단 규격 및 롤당 생산 원가] ---
    st.divider()
    st.subheader("📏 2. 원단 규격 및 생산 원가")
    
    c_w1, c_w2, c_w3 = st.columns(3)
    with c_w1:
        # 상단과 라벨이 같아도 key가 다르면 에러가 나지 않습니다.
        width_mm = st.number_input("비닐 폭 (mm)", value=630, key="final_width_mm") 
    with c_w2:
        length_m = st.number_input("원단 총 길이 (m)", value=1800, key="final_length_m") 
    with c_w3:
        thick_mm = st.number_input("비닐 두께 (mm)", value=0.009, step=0.001, format="%.3f", key="final_thick_mm") 

    # 무게 및 원가 계산
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
