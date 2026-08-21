def show_production_order():
    st.title("📋 생산 작업 지시서 생성")
    
    # 데이터 로드
    sheet_id = "13ldXPSVT7CFyNZRj-6Rlv3aXMqOhflquUtcZom5cJzU"
    df_products = pd.read_csv(f"https://docs.google.com/sheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote('상품목록')}")
    df_products.columns = [str(c).strip() for c in df_products.columns]
    
    # 입력 폼
    st.subheader("오늘의 생산 계획 입력")
    production_data = {}
    
    # 리스트 기반 입력 UI
    cols = st.columns(3)
    for idx, row in df_products.iterrows():
        sku = str(row['SKU ID'])
        name = row['상품명']
        col_idx = idx % 3
        production_data[sku] = cols[col_idx].number_input(f"{name} (수량)", min_value=0, value=0, key=f"prod_{sku}")

    if st.button("지시서 생성하기"):
        # 계산 로직
        order_list = []
        total_rolls = 0
        
        for sku, qty in production_data.items():
            if qty > 0:
                prod_row = df_products[df_products['SKU ID'].astype(str) == sku].iloc[0]
                # 롤당 생산 가능 수량 추출
                pcs_per_roll = clean_num(next((prod_row[k] for k in prod_row.index if '롤당' in k and '수량' in k), 1))
                needed_rolls = qty / pcs_per_roll if pcs_per_roll > 0 else 0
                
                order_list.append({
                    "상품명": prod_row['상품명'],
                    "목표수량": qty,
                    "필요롤수": round(needed_rolls, 2)
                })
                total_rolls += needed_rolls
        
        # 결과 출력 및 지시서 형태
        st.markdown("---")
        st.subheader("📄 작업 지시서")
        st.write(f"작성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        
        df_order = pd.DataFrame(order_list)
        st.table(df_order)
        st.metric("총 필요 원단(롤)", f"{round(total_rolls, 2)} 롤")
        
        # 프린트 버튼 (브라우저 인쇄 창 호출)
        st.button("🖨️ 인쇄하기", on_click=lambda: st.write("브라우저의 인쇄 기능(Ctrl+P)을 사용하세요."))
