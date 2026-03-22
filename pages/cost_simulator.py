import streamlit as st
from auth import check_password

# 1. 페이지 기본 설정 및 보안 체크 (반드시 최상단에 위치)
st.set_page_config(page_title="원가 시뮬레이터", page_icon="🏭", layout="wide")

if not check_password():
    st.stop()

def _section_fabric_spec() -> None:
    """섹션 1: 원단 규격 정밀 계산기"""
    st.subheader("📏 원단 규격 정밀 계산기")

    calc_mode = st.radio(
        "계산 모드",
        ["⚖️ 무게 산출", "🔍 두께 역산"],
        horizontal=True,
        key="cost_calc_mode",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        v_width = st.number_input("비닐 폭 (mm)", value=630, key="cost_v_width")
    with c2:
        v_length = st.number_input("원단 총 길이 (m)", value=1800, key="cost_v_length")

    if calc_mode == "⚖️ 무게 산출":
        with c3:
            v_thick = st.number_input(
                "두께 (mm)", value=0.009, format="%.3f", key="cost_v_thick"
            )
        weight = (v_width / 1000) * v_length * 2 * 0.92 * v_thick
        st.info(f"💡 예상 무게: {weight:.2f} kg")
    else:
        with c3:
            v_weight_in = st.number_input(
                "실제 무게 (kg)", value=13.8, key="cost_v_weight_in"
            )
        thick = v_weight_in / ((v_width / 1000) * v_length * 2 * 0.92)
        st.warning(f"💡 역산된 두께: {thick:.4f} mm")


def _section_material_price() -> float:
    """섹션 2: 원재료 혼합 단가 설정. 최종 단가(원/kg)를 반환합니다."""
    st.divider()
    st.subheader("🧪 1. 원재료 혼합 단가 설정")

    col1, col2 = st.columns(2)
    with col1:
        v_price = st.number_input("신원료 가격 (원/kg)", value=1530, key="mat_v_price")
        r_price = st.number_input("재생원료 가격 (원/kg)", value=1300, key="mat_r_price")
    with col2:
        v_ratio = st.slider("신원료 혼합 비율 (%)", 0, 100, 70, key="mat_v_ratio")
        st.caption(f"신원료 {v_ratio}% : 재생원료 {100 - v_ratio}%")

    st.write("---")
    col3, col4 = st.columns(2)
    with col3:
        c_price = st.number_input("조색제 가격 (원/kg)", value=2700, key="mat_c_price")
    with col4:
        c_ratio = st.number_input(
            "조색제 혼합 비율 (%)", value=2.5, step=0.1, format="%.1f", key="mat_c_ratio"
        )

    base_price = (v_price * (v_ratio / 100)) + (r_price * ((100 - v_ratio) / 100))
    final_unit_price = (base_price * (1 - c_ratio / 100)) + (c_price * (c_ratio / 100))
    st.success(f"🎨 **최종 원재료 단가: ₩{final_unit_price:,.2f} / kg**")
    return final_unit_price


def _section_production_cost(final_unit_price: float) -> None:
    """섹션 3: 원단 규격 및 롤당 생산 원가"""
    st.divider()
    st.subheader("📏 2. 원단 규격 및 생산 원가")

    c_w1, c_w2, c_w3 = st.columns(3)
    with c_w1:
        width_mm = st.number_input("비닐 폭 (mm)", value=630, key="prod_width_mm")
    with c_w2:
        length_m = st.number_input("원단 총 길이 (m)", value=1800, key="prod_length_m")
    with c_w3:
        thick_mm = st.number_input(
            "비닐 두께 (mm)", value=0.009, step=0.001, format="%.3f", key="prod_thick_mm"
        )

    total_weight = (width_mm / 1000) * length_m * 2 * 0.92 * thick_mm
    total_cost = total_weight * final_unit_price

    col_res1, col_res2 = st.columns(2)
    with col_res1:
        st.metric("예상 원단 무게", f"{total_weight:.2f} kg")
    with col_res2:
        st.metric("1롤당 제조 원가", f"₩{total_cost:,.0f}")


# 메인 실행부: render() 래퍼(wrapper)를 제거하고 직접 실행
st.title("🏭 원가 시뮬레이터")

_section_fabric_spec()
final_unit_price = _section_material_price()
_section_production_cost(final_unit_price)
