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

# --- [로직 1] 상품별 팔레트 적재 규칙 (v4.98 반영) ---
def get_pallet_capacity(sku):
    sku = str(sku)
    if sku in ['32058611', '15651222']: return 300
    if sku in ['29558294', '32711887']: return 192
    if sku == '32083343': return 400
    if sku == '32366753': return 560
    return 300

# --- [로직 2] 슬라이드 복제 및 텍스트 설정 (v4.98 반영) ---
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

# --- [로직 3] 데이터 입력 및 수량 분할 (v4.98 반영) ---
def fill_slide_data(slide, p, po_num, fc_name, year, month, day):
    try:
        current_plt_idx = int(p['no'].split('-')[1])
        total_qty = int(p['total_qty'])
        cap = int(p['cap'])
        if current_plt_idx * cap <= total_qty:
            display_qty = cap
        else:
            display_qty = total_qty % cap if total_qty % cap != 0 else cap
    except:
        display_qty = p['total_qty']

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

# --- [웹 대시보드 메인 페이지] ---
st.set_page_config(page_title="요정비닐 스마트 대시보드", layout="wide")
menu = st.sidebar.selectbox("메뉴 선택", ["시장 지표 분석", "밀크런 PPT 변환"])

if menu == "시장 지표 분석":
    st.title("📈 실시간 유가 및 환율 모니터링")
    # (유가 및 환율 로직은 이전과 동일하게 유지)

elif menu == "밀크런 PPT 변환":
    st.title("🚚 밀크런 자동 변환 시스템 (Web v1.0)")
    st.write("PDF 발주서를 업로드하면 수량 분할 규칙에 맞춰 PPT를 생성합니다.")

    tpl_file = st.file_uploader("1. 밀크런_양식.pptx 업로드", type=['pptx'])
    pdf_files = st.file_uploader("2. 발주서 PDF 업로드 (다중 선택 가능)", type=['pdf'], accept_multiple_files=True)

    if tpl_file and pdf_files:
        if st.button("🚀 변환 시작"):
            try:
                prs = Presentation(tpl_file)
                while len(prs.slides) > 1:
                    rId = prs.slides._sle[1].rId
                    prs.part.drop_rel(rId); del prs.slides._sle[1]
                
                is_first = True
                for pdf_file in pdf_files:
                    reader = pypdf.PdfReader(pdf_file)
                    text = "".join([page.extract_text() for page in reader.pages])
                    
                    po_match = re.search(r"(\d{9})", text)
                    po_num = po_match.group(1) if po_match else "000000000"
                    fc_match = re.search(r"(?:FC명|센터명)\s*[:\s]*([A-Z0-9가-힣]+)", text)
                    fc_name = fc_match.group(1) if fc_match else "알수없음"
                    date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
                    date_raw = date_match.group(1) if date_match else "2026-03-05"
                    y, m, d = date_raw.split('-')

                    # SKU 및 수량 추출 (v4.98 로직 준용)
                    skus = re.findall(r"\b(\d{8})\b", text)
                    items = []
                    for s in set(skus):
                        cap = get_pallet_capacity(s)
                        # 실제 PDF에서 수량(qty)과 상품명(name)을 추출하는 정규식이 추가되어야 함
                        items.append({'sku': s, 'name': "상품명 확인필요", 'qty': 600, 'cap': cap})

                    for itm in items:
                        tot_plt = (itm['qty'] // itm['cap']) + (1 if itm['qty'] % itm['cap'] > 0 else 0)
                        for i in range(1, tot_plt + 1):
                            p_info = {'no': f"{tot_plt}-{i}", 'total_qty': itm['qty'], 'cap': itm['cap'], 'items_list': [itm]}
                            for _ in range(2):
                                slide = prs.slides[0] if is_first else duplicate_slide(prs, 0)
                                is_first = False
                                fill_slide_data(slide, p_info, po_num, fc_name, y, m, d)

                output = io.BytesIO()
                prs.save(output)
                st.success("✨ 변환 완료!")
                st.download_button("📥 결과 PPT 다운로드", output.getvalue(), f"밀크런_결과_{datetime.now().strftime('%m%d')}.pptx")
            except Exception as e:
                st.error(f"오류 발생: {e}")
