import streamlit as st
import pandas as pd
import re
import io
from urllib.parse import quote

def show_reward_marketing():
    st.title("🎯 페어리테이블 리워드 마케팅 자동화 제어기 (V1.0)")
    st.markdown("---")

    # 1. 고마진 현물 리워드 원가 기준 세팅 (40L 기준)
    st.sidebar.header("🚀 마케팅 비용 설정")
    reward_mfg_cost = st.sidebar.number_input("📦 리워드 상품 제조원가 (원)", value=3500, step=100)
    reward_shipping = st.sidebar.number_input("🚚 리워드 발송 택배비 (원)", value=2400, step=100)
    total_reward_cost = reward_mfg_cost + reward_shipping
    
    st.sidebar.info(f"💡 리뷰 1건당 실질 마케팅 지출: {total_reward_cost:,}원")

    # 2. 신청 데이터 가상 로드 (실무에서는 구글 시트 API 연동 권장)
    st.subheader("📥 실시간 서포터즈 신청 현황")
    st.markdown("구글 폼이나 스트림릿 사용자 화면을 통해 들어온 신청 내역을 자동으로 검증합니다.")

    # 데모용 데이터프레임 구성 (실제 운영 시 시트 연동)
    demo_data = {
        "신청일시": ["2026-06-01 14:22", "2026-06-01 16:05", "2026-06-02 09:11", "2026-06-02 11:30"],
        "플랫폼": ["네이버 스마트스토어", "쿠팡", "네이버 스마트스토어", "네이버 스마트스토어"],
        "구매자명": ["김지은", "박상민", "이민지", "김지은"],
        "주문번호": ["20260601-9982312", "402-9918231-11203", "20260602-1102931", "20260601-9982312"],
        "포토리뷰링크": ["https://smartstore...", "https://coupang...", "https://smartstore...", "https://smartstore..."],
        "검증상태": ["대기", "대기", "대기", "대기"]
    }
    df_reward = pd.DataFrame(demo_data)

    # 3. 정보처리기사 스타일의 유효성 검증 알고리즘 알고리즘 (Data Validation)
    st.markdown("#### 🛡️ 알고리즘 자동 필터링 결과")
    
    verified_rows = []
    order_counts = df_reward["주문번호"].value_counts() # 중복 주문번호 체크용
    
    for idx, row in df_reward.iterrows():
        status = "✅ 검증완료"
        reason = "정상 신청"
        
        # Rule 1: 동일 주문번호로 중복 혜택을 노리는 체리피커 차단
        if order_counts[row["주문번호"]] > 1:
            status = "🚨 중복제외"
            reason = "동일한 주문번호로 다중 신청 감지"
            
        # Rule 2: 쿠팡/네이버 주문번호 정규식 자릿수 검증
        elif row["플랫폼"] == "쿠팡" and not re.match(r'^\d{3}-\d{7}-\d{7}$', str(row["주문번호"])):
            status = "⚠️ 양식오류"
            reason = "쿠팡 주문번호 포맷 불일치"
            
        elif row["플랫폼"] == "네이버 스마트스토어" and len(str(row["주문번호"])) < 10:
            status = "⚠️ 양식오류"
            reason = "네이버 주문번호 자릿수 부족"

        row["검증상태"] = f"{status} ({reason})"
        verified_rows.append(row)
        
    df_verified = pd.DataFrame(verified_rows)

    # 결과 대시보드 출력
    def highlight_reward(val):
        if "✅" in str(val): return 'background-color: #e8f5e9; color: #2e7d32; font-weight: bold;'
        if "🚨" in str(val): return 'background-color: #ffebee; color: #c62828; font-weight: bold;'
        if "⚠️" in str(val): return 'background-color: #fff8e1; color: #f57f17; font-weight: bold;'
        return ''

    st.dataframe(
        df_verified.style.map(highlight_reward, subset=["검증상태"]),
        use_container_width=True,
        hide_index=True
    )

    # 4. 액션 버튼: 엑셀 다운로드 (발송용)
    st.markdown("---")
    st.subheader("🚚 리워드 상품 발송 처리")
    
    # 정상 완료된 데이터만 추출하여 송장 뽑기 편하게 정렬
    df_shipping = df_verified[df_verified["검증상태"].str.contains("✅")].copy()
    df_shipping["발송상품"] = "40L 양면 비닐봉투 1세트(리워드)"
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="금일 발송 예정 리워드", value=f"{len(df_shipping)} 건")
    with col2:
        st.metric(label="예상 마케팅 소요 비용", value=f"{len(df_shipping) * total_reward_cost:,} 원")

    # 엑셀 다운로드 버퍼 생성
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_shipping.to_excel(writer, index=False, sheet_name='리워드발송명단')
        
    st.download_button(
        label="📥 발송 명단 엑셀 다운로드 (우체국/택배사 엑셀 업로드용)",
        data=buffer.getvalue(),
        file_name=f"페어리테이블_리워드발송_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.ms-excel"
    )

# 메인 파일 구동 조건에 추가 체킹
# show_reward_marketing()
