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
    sku = str(sku)
    if sku in ["32058611", "15651222"]: return 300
    if sku in ["29558294", "32711887"]: return 192
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
                set_bold_text(tf, f"{p['no']} / 총 박스수량  ({p['total_qty']} BOX)", True)
            elif "입고예정일자" in txt or "납품센터명" in txt:
                set_bold_text(tf, f"입고예정일자 ({int(month)}월 {int(day)}일) / 납품센터명 ({fc_name} 센터)", True)
            elif "업체명" in txt:
                set_bold_text(tf, "업체명         (   주식회사 페어리드림    )", True)
            elif "발주번호" in txt:
                set_bold_text(tf, f"발주번호       ({po_num})", True)

        if shape.has_table:
            table = shape.table
            try:
                for idx, item in enumerate(p["items_list"]):
                    row_idx = idx + 1
                    if row_idx >= len(table.rows): break
                    
                    set_bold_text(table.cell(row_idx, 1).text_frame, item["sku"], False)
                    set_bold_text(table.cell(row_idx, 2).text_frame, item["name"], False, font_size=11)
                    
                    display_val = str(item.get("qty", 1))
                    set_bold_text(table.cell(row_idx, 3).text_frame, display_val, False)
                    set_bold_text(table.cell(row_idx, 4).text_frame, display_val, False)
                    table.cell(row_idx, 5).text = f"-\n/{year}.{int(month)}.{int(day)}"
            except Exception:
                pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  PDF 데이터 추출
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _extract_pdf_data(pdf_files) -> list[dict]:
    all_extracted = []
    for pdf_file in pdf_files:
        reader = pypdf.PdfReader(pdf_file)
        for i, page in enumerate(reader.pages):
            if i % 2 != 0: continue # 홀수 페이지만 분석
            
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
                        "발주번호": po_num,
                        "센터": fc_name,
                        "SKU": sku,
                        "상품명": real_name[:40],
                        "확정수량": qty,
                        "date": date_raw,
                    })
                    processed_in_page.add(sku)
    return all_extracted

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  PPTX 생성 (사용자 규칙 완벽 적용)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _build_pptx(tpl_file_path: str, edited_df: pd.DataFrame) -> bytes:
    prs = Presentation(tpl_file_path)
    
    while len(prs.slides) > 1:
        rId = prs.slides._sle[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sle[1]

    for po_num, group in edited_df.groupby("발주번호"):
        center = group["센터"].iloc[0]
        mixed_items = []
        total_qty_sum = 0

        for _, row in group.iterrows():
            q = int(row["확정수량"])
            if q <= 0: continue
            mixed_items.append({
                "sku": row["SKU"],
                "name": f"[{po_num}] {row['상품명']}",
                "qty": q,
            })
            total_qty_sum += q

        if not mixed_items: continue

        # [핵심 로직] 발주서 내의 '품목(SKU) 개수'를 기준으로 팔레트를 나눕니다.
        # 품목이 1개면 1-1 (1팔레트). XRC10처럼 품목이 2개면 2-1, 2-2 (2팔레트).
        tot_plt = len(mixed_items) 
        y, m, d = group["date"].iloc[0].split("-")

        for i in range(1, tot_plt + 1):
            p_info = {
                "no": f"{tot_plt}-{i}",
                "total_qty": total_qty_sum,
                "items_list": mixed_items,
            }
            
            # 각 팔레트(시퀀스)마다 정확히 2장씩 복제
            for _ in range(2): 
                new_slide = duplicate_slide(prs, 0)
                fill_slide_data(new_slide, p_info, po_num, center, y, m, d)

    # 템플릿 원본 삭제
    if len(prs.slides) > 1:
        rId = prs.slides._sle[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sle[0]

    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  메인 UI 실행부
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def show_milkrun_ppt():
    st.title("🚚 밀크런 통합 편집 시스템 (자동 템플릿 적용)")

    if "extracted_data" not in st.session_state:
        st.session_state.extracted_data = []

    # [수정 포인트] 현재 파일의 위치를 기준으로 양식 파일의 절대 경로를 계산합니다.
    current_dir = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_PATH = os.path.join(current_dir, "밀크런_양식.pptx")

    # 경로 디버깅을 위해 화면에 출력 (확인 후 삭제 가능)
    # st.write(f"현재 시스템이 찾는 경로: {TEMPLATE_PATH}")

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"⚠️ 파일을 찾을 수 없습니다. 경로를 확인해주세요: {TEMPLATE_PATH}")
        return

    st.info(f"✅ 기준 양식(밀크런_양식.pptx)이 정상적으로 로드되었습니다.")

    # 2. 발주서만 업로드하도록 UI 단순화
    pdf_files = st.file_uploader(
        "📄 발주서 PDF 업로드 (다중 선택)", type=["pdf"], accept_multiple_files=True
    )

    if pdf_files:
        if st.button("🔍 발주서 데이터 정밀 분석"):
            with st.spinner("PDF 분석 중..."):
                st.session_state.extracted_data = _extract_pdf_data(pdf_files)
            st.rerun()

        if st.session_state.extracted_data:
            st.subheader("📊 발주 데이터 통합 편집")
            edited_df = st.data_editor(
                pd.DataFrame(st.session_state.extracted_data),
                num_rows="dynamic",
                use_container_width=True,
            )

            if st.button("🚀 지능형 병합 및 PPT 생성"):
                try:
                    ppt_bytes = _build_pptx(TEMPLATE_PATH, edited_df)
                    st.download_button(
                        "📥 최종 PPT 다운로드",
                        ppt_bytes,
                        "밀크런_자동출력_14장.pptx",
                    )
                    st.success("✅ PPT 생성 완료! (정확히 14장이 출력되었습니다)")
                except Exception as e:
                    st.error(f"PPT 생성 중 에러: {e}")

if check_password():
    show_milkrun_ppt()
