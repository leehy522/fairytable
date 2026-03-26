import copy
import io
import os
import re

import pandas as pd
import pypdf
import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from auth import check_password

# 1. 페이지 기본 설정
st.set_page_config(page_title="밀크런 PPT 자동변환", page_icon="🚚", layout="wide")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  헬퍼 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def get_pallet_capacity(sku: str) -> int:
    """SKU 번호에 따른 팔레트당 최대 적재 박스 수 (기존 로직 복구)"""
    sku = str(sku)
    if sku in ["32058611", "15651222"]: return 300  # 60리터 200매 등
    if sku in ["29558294", "32711887"]: return 192  # 100리터 등
    if sku == "32083343": return 400
    if sku == "32366753": return 560
    return 300

def duplicate_slide(prs: Presentation, index: int):
    template = prs.slides[index]
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(blank_layout)
    for shp in list(new_slide.shapes):
        new_slide.shapes._spTree.remove(shp.element)
    for shape in template.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
    return new_slide

def set_bold_text(text_frame, content, is_bold: bool = True, font_size=None) -> None:
    text_frame.text = str(content)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.bold = is_bold
            if font_size:
                run.font.size = Pt(font_size)

def fill_slide_data(slide, p: dict, po_num: str, fc_name: str, year: str, month: str, day: str) -> None:
    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            txt = shape.text

            if "박스수량" in txt or "BOX" in txt:
                # p['no']는 '2-1' 형태, p['item_total']은 해당 SKU의 총 수량
                set_bold_text(tf, f"{p['no']} / 총 박스수량  ({p['item_total']} BOX)", True)
            elif "입고예정일자" in txt or "납품센터명" in txt:
                set_bold_text(tf, f"입고예정일자 ({int(month)}월 {int(day)}일) / 납품센터명 ({fc_name} 센터)", True)
            elif "업체명" in txt:
                set_bold_text(tf, "업체명         (   주식회사 페어리드림    )", True)
            elif "발주번호" in txt:
                set_bold_text(tf, f"발주번호       ({po_num})", True)

        if shape.has_table:
            table = shape.table
            try:
                # SKU 단독 적재이므로 테이블의 첫 줄만 채움
                row_idx = 1
                set_bold_text(table.cell(row_idx, 1).text_frame, p["sku"], False)
                set_bold_text(table.cell(row_idx, 2).text_frame, p["name"], False, font_size=11)
                
                # 현재 팔레트에 담긴 수량
                display_val = str(p["current_qty"])
                set_bold_text(table.cell(row_idx, 3).text_frame, display_val, False)
                set_bold_text(table.cell(row_idx, 4).text_frame, display_val, False)
                table.cell(row_idx, 5).text = f"-\n/{year}.{int(month)}.{int(day)}"
            except Exception:
                pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  PDF 데이터 추출 및 PPT 생성 (SKU별 분리 로직 적용)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _extract_pdf_data(pdf_files) -> list[dict]:
    all_extracted = []
    for pdf_file in pdf_files:
        reader = pypdf.PdfReader(pdf_file)
        for i, page in enumerate(reader.pages):
            if i % 2 != 0: continue 
            text = page.extract_text() + "\n"
            po_match = re.search(r"(?:발주번호|PO|no|Info)\s*[:\s\n]*(\d{9})", text, re.I) or re.search(r"(\d{9})", text)
            if not po_match: continue
            po_num = po_match.group(1)
            fc_match = re.search(r"(?:FC명|FC\s*Name|센터명)\s*[:\s\n]*([A-Z0-9가-힣]+)", text, re.I) or re.search(r"([가-힣]+)센터", text)
            fc_name = fc_match.group(1).strip() if fc_match else "알수없음"
            date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
            date_raw = date_match.group(1) if date_match else "2026-03-27"

            processed_in_page = set()
            for m in re.finditer(r"\b(\d{8})\b", text):
                sku = m.group(1)
                if sku in processed_in_page: continue
                block = text[m.end():m.end() + 450]
                name_search = re.search(r"([가-힣]{2,}[가-힣\s\d\-\(\)]+)", block)
                real_name = name_search.group(1).strip() if name_search else "상품명확인"
                nums = re.findall(r"\b\d{1,4}\b", block)
                qty = int(nums[1]) if len(nums) >= 2 else (int(nums[0]) if len(nums) == 1 else 0)
                if qty > 0:
                    all_extracted.append({
                        "발주번호": po_num, "센터": fc_name, "SKU": sku,
                        "상품명": real_name[:40], "확정수량": qty, "date": date_raw,
                    })
                    processed_in_page.add(sku)
    return all_extracted

def _build_pptx(tpl_file_path: str, edited_df: pd.DataFrame) -> bytes:
    prs = Presentation(tpl_file_path)
    
    # 첫 번째 슬라이드만 남기고 모두 삭제
    slide_ids = [s.slide_id for s in prs.slides]
    for s_id in slide_ids[1:]:
        slide_layout_index = prs.slides._sldIdLst.index(next(s for s in prs.slides._sldIdLst if s.id == s_id))
        del prs.slides._sldIdLst[slide_layout_index]

    # 발주번호별로 먼저 묶음
    for po_num, group in edited_df.groupby("발주번호"):
        center = group["센터"].iloc[0]
        y, m, d = group["date"].iloc[0].split("-")
        
        # [핵심] 발주서 내의 각 SKU별로 루프를 돌려 팔레트를 생성 (혼적 방지)
        for _, row in group.iterrows():
            sku_code = row["SKU"]
            sku_name = f"[{po_num}] {row['상품명']}"
            qty_total = int(row["확정수량"])
            cap = get_pallet_capacity(sku_code)
            
            # 해당 SKU의 팔레트 수 계산
            tot_plt = (qty_total // cap) + (1 if qty_total % cap > 0 else 0)
            
            for i in range(1, tot_plt + 1):
                # 현재 팔레트에 담길 수량 결정
                if i * cap <= qty_total:
                    current_plt_qty = cap
                else:
                    current_plt_qty = qty_total % cap if qty_total % cap != 0 else cap
                
                p_info = {
                    "no": f"{tot_plt}-{i}",
                    "total_qty": qty_total, # 이 SKU의 총 수량
                    "item_total": qty_total,
                    "sku": sku_code,
                    "name": sku_name,
                    "current_qty": current_plt_qty
                }
                
                for _ in range(2): 
                    new_slide = duplicate_slide(prs, 0)
                    fill_slide_data(new_slide, p_info, po_num, center, y, m, d)

    # 원본 템플릿 삭제
    del prs.slides._sldIdLst[0]

    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  메인 UI 실행부
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def show_milkrun_ppt():
    st.title("🚚 SKU 단독 적재 밀크런 시스템")

    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = []

    current_dir = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_PATH = os.path.join(current_dir, "밀크런_양식.pptx")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"⚠️ 양식 파일이 없습니다: {TEMPLATE_PATH}")
        return

    st.success(f"✅ 양식 로드 완료")
    pdf_files = st.file_uploader("📄 발주서 PDF 업로드", type=["pdf"], accept_multiple_files=True)

    if pdf_files:
        if st.button("🔍 데이터 분석"):
            st.session_state.extracted_data = _extract_pdf_data(pdf_files)
            st.rerun()

        if st.session_state.extracted_data:
            edited_df = st.data_editor(pd.DataFrame(st.session_state.extracted_data), use_container_width=True)

            if st.button("🚀 PPT 생성"):
                try:
                    ppt_bytes = _build_pptx(TEMPLATE_PATH, edited_df)
                    st.download_button("📥 다운로드", ppt_bytes, "밀크런_단독적재_결과.pptx")
                except Exception as e:
                    st.error(f"오류 발생: {e}")

if check_password():
    show_milkrun_ppt()
