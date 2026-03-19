"""
pages/milkrun_ppt.py — 🚚 밀크런 PPT 변환 (v4.98 로직)
PDF 발주서를 파싱 → 편집 → PPTX 생성까지 담당합니다.

내부 헬퍼 함수:
  get_pallet_capacity  : SKU별 팔레트 적재량 반환
  duplicate_slide      : 슬라이드 복제
  set_bold_text        : TextFrame 굵기/폰트 설정
  fill_slide_data      : 슬라이드에 데이터 채우기
  _extract_pdf_data    : PDF → 발주 데이터 추출
  _build_pptx          : 편집된 DataFrame → PPTX 생성
"""

import copy
import io
import re

import pandas as pd
import pypdf
import streamlit as st
from pptx import Presentation
from pptx.util import Pt


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  헬퍼 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def get_pallet_capacity(sku: str) -> int:
    """SKU 번호에 따라 팔레트 최대 적재 박스 수를 반환합니다."""
    sku = str(sku)
    if sku in ["32058611", "15651222"]:
        return 300
    if sku in ["29558294", "32711887"]:
        return 192
    if sku == "32083343":
        return 400
    if sku == "32366753":
        return 560
    return 300


def duplicate_slide(prs: Presentation, index: int):
    """index번 슬라이드를 맨 뒤에 복제해서 반환합니다."""
    template = prs.slides[index]
    blank_layout = (
        prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    )
    new_slide = prs.slides.add_slide(blank_layout)
    # 빈 레이아웃의 기본 도형 제거
    for shp in list(new_slide.shapes):
        new_slide.shapes._spTree.remove(shp.element)
    # 원본 도형 복사
    for shape in template.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
    return new_slide


def set_bold_text(text_frame, content, is_bold: bool = True, font_size=None) -> None:
    """TextFrame 전체를 content로 교체하고 굵기/크기를 설정합니다."""
    text_frame.text = str(content)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.bold = is_bold
            if font_size:
                run.font.size = Pt(font_size)


def fill_slide_data(slide, p: dict, po_num: str, fc_name: str,
                    year: str, month: str, day: str) -> None:
    """슬라이드의 텍스트/표를 발주 데이터로 채웁니다."""
    try:
        current_plt_idx = int(p["no"].split("-")[1])
        total_qty = int(p["total_qty"])
        cap = int(p["cap"])
        display_qty = (
            cap
            if current_plt_idx * cap <= total_qty
            else (total_qty % cap if total_qty % cap != 0 else cap)
        )
    except Exception:
        display_qty = p["total_qty"]

    for shape in slide.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            txt = shape.text

            if "박스수량" in txt or "BOX" in txt:
                set_bold_text(
                    tf, f"{p['no']} / 총 박스수량  ({p['total_qty']} BOX)", True
                )
            elif "입고예정일자" in txt or "납품센터명" in txt:
                set_bold_text(
                    tf,
                    f"입고예정일자 ({int(month)}월 {int(day)}일) / 납품센터명 ({fc_name} 센터)",
                    True,
                )
            elif "업체명" in txt:
                set_bold_text(tf, "업체명         (   주식회사 페어리드림    )", True)
            elif "발주번호" in txt:
                set_bold_text(tf, f"발주번호       ({po_num})", True)

        if shape.has_table:
            table = shape.table
            try:
                item_count = len(p["items_list"])
                for idx, item in enumerate(p["items_list"]):
                    row_idx = idx + 1
                    if row_idx >= len(table.rows):
                        break
                    set_bold_text(table.cell(row_idx, 1).text_frame, item["sku"], False)
                    set_bold_text(
                        table.cell(row_idx, 2).text_frame, item["name"], False, font_size=11
                    )
                    display_val = (
                        str(item.get("qty", item.get("확정수량", 1)))
                        if item_count > 1
                        else str(item.get("cap", p.get("cap", 300)))
                    )
                    set_bold_text(table.cell(row_idx, 3).text_frame, display_val, False)
                    set_bold_text(table.cell(row_idx, 4).text_frame, display_val, False)
                    table.cell(row_idx, 5).text = f"-\n/{year}.{int(month)}.{int(day)}"
            except Exception:
                pass


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  PDF 데이터 추출
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _extract_pdf_data(pdf_files) -> list[dict]:
    """업로드된 PDF 리스트에서 발주 데이터를 추출합니다. (홀수 페이지만)"""
    all_extracted = []

    for pdf_file in pdf_files:
        reader = pypdf.PdfReader(pdf_file)
        for i, page in enumerate(reader.pages):
            if i % 2 != 0:          # 짝수 인덱스(짝수 장) 스킵
                continue

            text = page.extract_text() + "\n"

            po_match = re.search(
                r"(?:발주번호|PO|no|Info)\s*[:\s\n]*(\d{9})", text, re.I
            ) or re.search(r"(\d{9})", text)
            if not po_match:
                continue

            po_num = po_match.group(1)
            fc_match = re.search(
                r"(?:FC명|FC\s*Name|센터명)\s*[:\s\n]*([A-Z0-9가-힣]+)", text, re.I
            ) or re.search(r"([가-힣]+)센터", text)
            fc_name = fc_match.group(1).strip() if fc_match else "알수없음"

            date_match = re.search(r"(\d{4}-\d{2}-\d{2})", text)
            date_raw = date_match.group(1) if date_match else "2026-03-12"

            processed_in_page: set = set()
            for m in re.finditer(r"\b(\d{8})\b", text):
                sku = m.group(1)
                if sku in processed_in_page:
                    continue
                cap = get_pallet_capacity(sku)
                block = text[m.end():m.end() + 450]
                name_search = re.search(r"([가-힣]{2,}[가-힣\s\d\-\(\)]+)", block)
                real_name = name_search.group(1).strip() if name_search else "상품명확인"
                nums = re.findall(r"\b\d{1,4}\b", block)
                qty = (
                    int(nums[1]) if len(nums) >= 2
                    else (int(nums[0]) if len(nums) == 1 else 0)
                )
                if qty > 0:
                    all_extracted.append({
                        "발주번호": po_num,
                        "센터": fc_name,
                        "SKU": sku,
                        "상품명": real_name[:40],
                        "확정수량": qty,
                        "적재량": cap,
                        "date": date_raw,
                    })
                    processed_in_page.add(sku)

    return all_extracted


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  PPTX 생성
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _build_pptx(tpl_file, edited_df: pd.DataFrame) -> bytes:
    """편집된 DataFrame으로 최종 PPTX를 생성하고 bytes를 반환합니다."""
    prs = Presentation(tpl_file)

    # 슬라이드를 1장만 남기기
    while len(prs.slides) > 1:
        rId = prs.slides._sle[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sle[1]

    is_first = True
    for center, group in edited_df.groupby("센터"):
        po_list = sorted([str(p) for p in group["발주번호"].unique()])
        all_pos = ", ".join(po_list)

        mixed_items = []
        total_qty_sum = 0

        for _, row in group.iterrows():
            q = int(row["확정수량"])
            if q <= 0:
                continue
            mixed_items.append({
                "sku": row["SKU"],
                "name": f"[{row['발주번호']}] {row['상품명']}",
                "qty": q,
            })
            total_qty_sum += q

        if not mixed_items:
            continue

        cap = int(group["적재량"].iloc[0])
        tot_plt = (
            1
            if total_qty_sum < 300
            else (total_qty_sum // cap) + (1 if total_qty_sum % cap > 0 else 0)
        )
        y, m, d = group["date"].iloc[0].split("-")

        for i in range(1, tot_plt + 1):
            p_info = {
                "no": f"{tot_plt}-{i}",
                "total_qty": total_qty_sum,
                "cap": cap,
                "items_list": mixed_items,
            }
            for _ in range(2):    # 슬라이드 2장씩 (앞/뒤)
                if is_first:
                    slide = prs.slides[0]
                    is_first = False
                else:
                    slide = duplicate_slide(prs, 0)
                fill_slide_data(slide, p_info, all_pos, center, y, m, d)

    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Streamlit 렌더링
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def render() -> None:
    st.title("🚚 밀크런 통합 편집 시스템 (v4.98 로직 이식)")

    tpl_file = st.file_uploader("1. 양식 PPT 업로드", type=["pptx"])
    pdf_files = st.file_uploader(
        "2. 발주서 PDF 업로드 (다중 선택)", type=["pdf"], accept_multiple_files=True
    )

    if not (tpl_file and pdf_files):
        return

    # ── PDF 분석 ──────────────────────────────────────
    if st.button("🔍 발주서 데이터 정밀 분석 (홀수 장 전용)"):
        with st.spinner("PDF 분석 중..."):
            st.session_state.extracted_data = _extract_pdf_data(pdf_files)
        st.rerun()

    # ── 편집 테이블 ───────────────────────────────────
    if not st.session_state.get("extracted_data"):
        return

    st.subheader("📊 발주 데이터 통합 편집")
    edited_df = st.data_editor(
        pd.DataFrame(st.session_state.extracted_data),
        num_rows="dynamic",
        use_container_width=True,
    )

    # ── PPT 생성 ──────────────────────────────────────
    if st.button("🚀 지능형 합짐 및 PPT 생성"):
        try:
            ppt_bytes = _build_pptx(tpl_file, edited_df)
            st.download_button(
                "📥 최종 PPT 다운로드",
                ppt_bytes,
                "밀크런_수량수정_결과.pptx",
            )
            st.success("✅ PPT 생성 완료!")
        except Exception as e:
            st.error(f"PPT 생성 중 에러: {e}")