# ... (앞부분의 모든 import 및 헬퍼 함수들은 그대로 유지하세요) ...

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  Streamlit 렌더링
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def render() -> None:
    st.title("🚚 밀크런 통합 편집 시스템 (v4.98 로직 이식)")
    st.info("💡 PDF 발주서의 홀수 페이지만 분석하여 PPT 양식에 채워 넣습니다.")

    tpl_file = st.file_uploader("1. 양식 PPT 업로드", type=["pptx"])
    pdf_files = st.file_uploader(
        "2. 발주서 PDF 업로드 (다중 선택)", type=["pdf"], accept_multiple_files=True
    )

    if not (tpl_file and pdf_files):
        # 데이터가 없을 때 세션 초기화 버튼 제공
        if st.sidebar.button("🧹 분석 데이터 초기화"):
            if "extracted_data" in st.session_state:
                del st.session_state.extracted_data
            st.rerun()
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
    # 사용자가 데이터를 수정할 수 있는 표 제공
    edited_df = st.data_editor(
        pd.DataFrame(st.session_state.extracted_data),
        num_rows="dynamic",
        use_container_width=True,
        key="milkrun_editor"
    )

    # ── PPT 생성 ──────────────────────────────────────
    st.divider()
    if st.button("🚀 지능형 합짐 및 PPT 생성"):
        try:
            with st.spinner("PPT 파일을 생성하고 있습니다..."):
                ppt_bytes = _build_pptx(tpl_file, edited_df)
                st.download_button(
                    "📥 최종 PPT 다운로드",
                    ppt_bytes,
                    "밀크런_수량수정_결과.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.success("✅ PPT 생성 완료! 위 버튼을 눌러 다운로드하세요.")
        except Exception as e:
            st.error(f"PPT 생성 중 에러가 발생했습니다: {e}")

