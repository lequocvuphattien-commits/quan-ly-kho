# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất":
    st.header("Nhập/Xuất hàng loạt")
    if 'cart' not in st.session_state: st.session_state.cart = []
    
    products = get_cached_products(service)
    if products:
        # Tạo map tồn kho để kiểm tra nhanh
        stock_map = {str(p[1]).strip(): float(p[4]) for p in products}
        p_dict = {f"{p[1]} - {p[2]}": p[1] for p in products}
        
        col1, col2, col3 = st.columns([2, 1, 1])
        sel = col1.selectbox("Chọn hàng", list(p_dict.keys()))
        qty = col2.number_input("Số lượng", min_value=1.0, step=1.0)
        typ = col3.radio("Loại", ["Nhập", "Xuất"], horizontal=True)
        
        if st.button("➕ Thêm vào lưới"):
            prod_code = p_dict[sel]
            # Kiểm tra tồn kho trước khi thêm vào lưới
            if typ == "Xuất":
                current_stock = stock_map.get(str(prod_code).strip(), 0)
                # Tính lượng đã xuất trong lưới chờ
                cart_df = pd.DataFrame(st.session_state.cart) if st.session_state.cart else pd.DataFrame(columns=["Mã", "Loại", "Số lượng"])
                out_in_cart = cart_df[(cart_df["Mã"] == prod_code) & (cart_df["Loại"] == "Xuất")]["Số lượng"].sum()
                
                if qty + out_in_cart > current_stock:
                    st.error(f"❌ Không đủ tồn kho! (Tồn hiện tại: {current_stock})")
                else:
                    st.session_state.cart.append({"Mã": prod_code, "Loại": typ, "Số lượng": qty})
                    st.rerun()
            else:
                st.session_state.cart.append({"Mã": prod_code, "Loại": typ, "Số lượng": qty})
                st.rerun()

        if st.session_state.cart:
            st.write("📋 Lưới chờ:", pd.DataFrame(st.session_state.cart))
            if st.button("✅ Xác nhận tất cả"):
                for item in st.session_state.cart:
                    service.add_transaction(item["Mã"], item["Số lượng"], item["Loại"], "Hàng loạt")
                    service.update_stock(item["Mã"], item["Số lượng"], item["Loại"])
                st.session_state.cart = []
                clear_all_caches(); st.success("Hoàn tất!"); st.rerun()
            if st.button("🗑️ Hủy"):
                st.session_state.cart = []; st.rerun()