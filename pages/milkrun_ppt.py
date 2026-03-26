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
    if sku in ["32058611", "15651222"]: return 300  # 60리터 200매
    if sku in ["29558294", "32711887"]: return 192  # 100리터
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
                # p['no']는 '2-1' 형태, p['item_total']은 해당 SKU의 박스수량
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
                row_idx = 1
                set_bold_text(table.cell(row_idx, 1).text_frame, p["sku"], False)
                set_bold_text(table.cell(row_idx, 2).text_frame, p["name"], False, font_size=11)
                set_bold_text(table.cell(row_idx, 3).text_frame, str(p["current_qty"]), False)
                set_bold_text(table.cell(row_idx, 4).text_frame, str(p["current_qty"]), False)
                table.cell(row_idx, 5).text = f"-\n/{year}.{int(month)}.{int(day)}"
            except Exception: pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  메인 로직 (센터별 통합 시퀀스 계산)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _build_pptx(tpl_file_path: str, edited_df: pd.DataFrame) -> bytes:
    prs = Presentation(tpl_file_path)
    
    # 템플릿 외 슬라이드 삭제
    slide_ids = [s.slide_id for s in prs.slides]
    for s_id in slide_ids[1:]:
        idx = prs.slides._sldIdLst.index(next(s for s in prs.slides._sldIdLst if s.id == s_id))
        del prs.slides._sldIdLst[idx]

    for po_num, group in edited_df.groupby("발주번호"):
        center = group["센터"].iloc[0]
        y, m, d = group["date"].iloc[0].split("-")
        
        # 1. 이 발주서(센터)에서 필요한 '총 팔레트 수'를 먼저 계산
        total_pallets_in_po = 0
        pallet_plan = [] # 각 SKU별 팔레트 할당 계획 저장
        
        for _, row in group.iterrows():
            qty = int(row["확정수량"])
            cap = get_pallet_capacity(row["SKU"])
            num_plt = (qty // cap) + (1 if qty % cap > 0 else 0)
            total_pallets_in_po += num_plt
            pallet_plan.append((row, num_plt, cap))

        # 2. 통합 시퀀스(N-M) 적용하여 슬라이드 생성
        global_plt_idx = 1
        for row, num_plt, cap in pallet_plan:
            qty_total = int(row["확정수량"])
            
            for i in range(1, num_plt + 1):
                # 현재 팔레트 적재량 계산
                if i * cap <= qty_total:
                    curr_qty = cap
                else:
                    curr_qty = qty_total % cap if qty_total % cap != 0 else cap
                
                p_info = {
                    "no": f"{total_pallets_in_po}-{global_plt_idx}", # 예: 2-1, 2-2
                    "item_total": qty_total,
                    "sku": row["SKU"],
                    "name": f"[{po_num}] {row['상품명']}",
                    "current_qty": curr_qty
                }
                
                for _ in range(2): 
                    new_slide = duplicate_slide(prs, 0)
                    fill_slide_data(new_slide, p_info, po_num, center, y, m, d)
                
                global_plt_idx += 1

    del prs.slides._sldIdLst[0]
    ppt_out = io.BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()

# ... (추출 및 UI 부분은 이전과 동일하므로 생략, build_pptx 호출부만 확인)
