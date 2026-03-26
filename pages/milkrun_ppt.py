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
#  헬퍼 함수 (적재 용량 및 슬라이드 조작)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def get_pallet_capacity(sku: str) -> int:
    """SKU 번호에 따른 팔레트당 최대 적재 박스 수"""
    sku = str(sku)
    # 60리터 200매 등 (예: 15651222)
    if sku in ["32058611", "15651222"]: return 300  
    # 100리터 등 (예: 29558294)
    if sku in ["29558294", "32711887"]: return 192  
    if sku == "32083343": return 400
    if sku == "32366753": return 560
    return 300

def duplicate_slide(prs: Presentation, index: int):
    """지정한 인덱스의 슬라이드를 복제하여 맨 뒤에 추가"""
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
    """텍스트 프레임에 내용을 채우고 스타일 적용"""
    text_frame.text = str(content)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.bold = is_bold
            if font_size:
                run.font.size = Pt(font_size)

def fill_slide_data(slide, p: dict, po_num: str, fc_name: str, year: str, month: str, day: str) -> None:
    """슬라이드 내 텍스트 및 테이블 업데이트"""
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
                row_idx = 1 # SKU별 단독 적재이므로 1번 행만 사용
                set_bold_text(table.cell(row_idx, 1).text_frame, p["sku"], False)
                set_bold_text(table.cell(row_idx, 2).text_frame, p["name"], False, font_size=11)
                set_bold_text(table.cell(row_idx, 3).text_frame, str(p["current_qty"]), False)
                set_bold_text(table.cell(row_idx, 4).text_frame, str(p["current_qty"]), False)
                table.cell(row_idx, 5).text = f"-\n/{year}.{int(month)}.{int(day)}"
            except Exception: pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  비즈니스 로직 (PDF 추출 및 PPT 빌드)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _extract_pdf_data(pdf_files) -> list[dict]:
    """PDF에서 발주 정보를 정밀 추출"""
    all_extracted = []
    for pdf_file in pdf_files:
        reader = pypdf.PdfReader(pdf_file)
        for i, page in enumerate(reader.pages):
            if i % 2 != 0: continue # 쿠팡 발주서는 홀수 페이지만 데이터 존재
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
    """템플릿을 사용하여 센터별 통합 시퀀스가 적용된 PPT 생성"""
    prs = Presentation(tpl_file_path)
    
    # 템플릿 외 기존 슬라이드 정리
    slide_ids = [s.slide_id for s in prs.slides]
    for s_id in slide_ids[1:]:
        idx = prs.slides._sldIdLst.index(next(s for s in prs.slides._sldIdLst if s.id == s_id))
        del prs.slides._sldIdLst[idx]

    # 발주번호(센터)별 그룹화 처리
    for po_num, group in edited_df.groupby("발주번호"):
        center = group["센터"].iloc[0]
        y, m, d = group["date"].iloc[0].split("-")
        
        # 1. 해당 센터의 총 팔레트 수 사전 계산 (N-M의 'N' 결정)
        total_pallets_in_po = 0
        pallet_plan = [] 
        
        for _, row in group.iterrows():
            qty = int(row["확정수량"])
            cap = get_pallet_capacity(row["SKU"])
            num_plt = (qty // cap) + (1 if qty % cap > 0 else 0)
            total_pallets_in_po += num_plt
            pallet_plan.append((row, num_plt, cap))

        # 2. 통합 시퀀스 적용 및 슬라이드 복제 (N-M의 'M' 부여)
        global_plt_idx = 1
        for row, num_plt, cap in pallet_plan:
            qty_total = int(row["확정수량"])
            sku_code = row["SKU"]
            sku_name = f"[{po_num}] {row['상품명']}"
            
            for i in range(1, num_plt + 1):
                # 팔레트별 적재 수량 분배
                if i * cap <= qty_total:
                    curr_qty = cap
                else:
                    curr_qty = qty_total % cap if qty_total % cap != 0 else cap
                
                p_info = {
                    "no": f"{total_pallets_in_po}-{global_plt_idx}", # 통합 번호
                    "item_total": qty_total,
                    "sku": sku_code,
                    "name": sku_name,
                    "current_qty": curr_qty
                }
                
                # 제출용/보관용 총 2장씩 생성
                for _ in range(2): 
                    new_slide = duplicate_slide(prs, 0)
                    fill_slide_data(new_slide, p_info, po_num, center, y, m, d)
                
                global_plt_idx += 1

    # 기준이 되었던 0번 템플릿 슬라이드 삭제
    del prs.slides._sldIdLst[0]
    
    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  메인 UI 실행부
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def show_milkrun_ppt():
    st.title("🚚 페어리테림 밀크런 통합 관리 시스템")

    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = []

    # 자동 경로 설정 (milkrun_ppt.py와 같은 폴더의 '밀크런_양식.pptx' 로드)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_PATH = os.path.join(current_dir, "밀크런_양식.pptx")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"⚠️ 기준 양식 파일이 없습니다: {TEMPLATE_PATH}")
        return

    st.success(f"✅ '밀크런_양식.pptx' 로드 완료")
    pdf_files = st.file_uploader("📄 발주서 PDF 업로드 (다중 선택 가능)", type=["pdf"], accept_multiple_files=True)

    if pdf_files:
        if st.button("🔍 발주 정보 정밀 분석"):
            with st.spinner("PDF 데이터 파싱 중..."):
                st.session_state.extracted_data = _extract_pdf_data(pdf_files)
            st.rerun()

        if st.session_state.extracted_data:
            st.subheader("📊 발주 데이터 통합 편집 (수정 가능)")
            edited_df = st.data_editor(pd.DataFrame(st.session_state.extracted_data), use_container_width=True)

            if st.button("🚀 PPT 통합 생성 (SKU별 분리 적용)"):
                try:
                    ppt_bytes = _build_pptx(TEMPLATE_PATH, edited_df)
                    st.download_button("📥 최종 PPT 결과물 다운로드", ppt_bytes, "밀크런_결과_14장.pptx")
                    st.success("✅ PPT 생성이 완료되었습니다!")
                except Exception as e:
                    st.error(f"작업 중 오류 발생: {e}")

if check_password():
    show_milkrun_ppt()
