import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report

# Khởi tạo dịch vụ
service = DataService(mode="ONLINE")

# --- BỘ NHỚ ĐỆM (CACHE) ---
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc):
    return _svc.get_products()

@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc):
    return _svc.get_history()

# Cấu hình trang
st.set_page_config(page_title="Quản Lý Kho Hàng", layout="wide")
st.title("📦 Quản lý kho hàng")

menu = st.sidebar.selectbox("Menu", ["Danh mục hàng hóa", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 1: DANH MỤC HÀNG HÓA ---
if menu == "Danh mục hàng hóa":
    st.header("Danh mục hàng hóa")
    products = get_cached_products(service)
    
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        
        # Hiển thị bảng
        st.dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]], width='stretch', hide_index=True)
        
        # Xuất Excel chuyên nghiệp
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df[["Mã", "Tên", "Đvt", "Tồn"]].to_excel(writer, index=False, sheet_name='DanhMuc')
            worksheet = writer.sheets['DanhMuc']
            # Cố định tiêu đề
            worksheet.freeze_panes = 'A2'
            # Thêm Filter
            max_row = worksheet.max_row
            max_col = worksheet.max_column
            worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
            # Chỉnh độ rộng cột
            for col in range(1, max_col + 1):
                worksheet.column_dimensions[get_column_letter(col)].width = 15
        
        st.download_button(
            label="📥 Xuất danh mục ra Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="DanhMucHangHoa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with st.form("add_form", clear_on_submit=True):
        st.subheader("Thêm hàng hóa mới")
        code, name, unit = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
        if st.form_submit_button("Thêm hàng hóa"):
            if not code or not name: st.warning("Nhập đủ Mã và Tên!")
            elif service.check_product_exists(code.upper()): st.error("Mã đã tồn tại!")
            else:
                service.add_product(code, name, unit)
                st.cache_data.clear()
                st.success("Đã thêm thành công!"); st.rerun()

# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất":
    st.header("🔄 Nhập/Xuất hàng hóa")
    
    # Khởi tạo giỏ hàng trong bộ nhớ tạm (Session State)
    if 'cart' not in st.session_state: 
        st.session_state.cart = []
    
    products = get_cached_products(service)
    if products:
        # Tạo dictionary chứa TẤT CẢ thông tin để tra cứu nhanh (Mã, Tên, Đvt, Tồn)
        p_dict = {f"{p['Mã']} - {p['Tên']}": p for p in products}
        
        # --- KHUNG NHẬP LIỆU GIAO DỊCH ---
        with st.container(border=True):
            # 1. Loại giao dịch (Nút Button ngang)
            typ = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
            
            # 2. Chọn hàng (Listbox có hỗ trợ gõ tìm kiếm ký tự đại diện mặc định của Streamlit)
            sel = st.selectbox(
                "Chọn hàng hóa", 
                options=list(p_dict.keys()), 
                index=None, 
                placeholder="🔍 Gõ tìm kiếm mã hoặc tên hàng..."
            )
            
            # Chia cột cho Số lượng, Ghi chú và Tồn kho
            c1, c2, c3 = st.columns([1, 1.5, 1])
            with c1:
                # 3. Số lượng: Để ô trống (value=None), hiển thị format có dấu phẩy
                qty = st.number_input("Số lượng", min_value=1, value=None, step=1, format="%d", placeholder="Nhập số...")
            with c2:
                # 4. Ghi chú (để trống mặc định)
                note = st.text_input("Ghi chú", placeholder="Nhập ghi chú...")
            with c3:
                # 5. Tồn kho động
                if sel:
                    current_stock = float(p_dict[sel]['Tồn'])
                    unit = p_dict[sel]['Đvt']
                    # Hiển thị số tồn định dạng có dấu phẩy (VD: 2,055)
                    st.markdown(f"<div style='padding-top: 30px; font-size: 18px; font-weight: bold; color: #28a745;'>Tồn: {current_stock:,.0f} {unit}</div>", unsafe_allow_html=True)
                else:
                    st.markdown("<div style='padding-top: 30px; font-size: 18px;'>Tồn: --</div>", unsafe_allow_html=True)

            # Nút thêm vào lưới
            if st.button("➕ Thêm vào lưới"):
                if not sel or not qty:
                    st.warning("⚠️ Vui lòng chọn hàng hóa và nhập số lượng!")
                else:
                    prod_data = p_dict[sel]
                    prod_code = prod_data["Mã"]
                    current_stock = float(prod_data["Tồn"])

                    # Logic kiểm tra tồn kho xuất
                    if typ == "Xuất":
                        cart_df = pd.DataFrame(st.session_state.cart) if st.session_state.cart else pd.DataFrame(columns=["Mã", "Loại", "Số lượng"])
                        out_in_cart = cart_df[(cart_df["Mã"] == prod_code) & (cart_df["Loại"] == "Xuất")]["Số lượng"].sum() if not cart_df.empty else 0
                        if qty + out_in_cart > current_stock:
                            st.error(f"❌ Không đủ tồn kho! (Tồn hiện tại: {current_stock:,.0f})")
                            st.stop() # Chặn không cho thêm vào lưới

                    # Thêm vào giỏ hàng
                    st.session_state.cart.append({
                        "Mã": prod_code,
                        "Tên HH": prod_data["Tên"],
                        "Đvt": prod_data["Đvt"],
                        "Loại": typ,
                        "Số lượng": qty,
                        "Ghi chú": note if note else ""
                    })
                    st.rerun() # Rerun cực nhanh vì không bị xóa cache

        # --- HIỂN THỊ LƯỚI CHỜ (CÓ THỂ CHỈNH SỬA) ---
        if st.session_state.cart:
            st.markdown("### 📋 Lưới chờ xử lý")
            st.caption("💡 *Mẹo: Click đúp chuột vào ô Số lượng hoặc Ghi chú để sửa trực tiếp trên bảng.*")
            
            # Sử dụng data_editor để biến bảng thành Excel thu nhỏ
            edited_df = st.data_editor(
                pd.DataFrame(st.session_state.cart),
                column_config={
                    "Mã": st.column_config.TextColumn("Mã HH", disabled=True),
                    "Tên HH": st.column_config.TextColumn("Tên HH", disabled=True),
                    "Đvt": st.column_config.TextColumn("Đvt", disabled=True),
                    "Loại": st.column_config.TextColumn("Loại", disabled=True),
                    # Cột số lượng cho phép sửa, định dạng phân cách hàng nghìn
                    "Số lượng": st.column_config.NumberColumn("Số lượng", required=True, min_value=1, format="%d"),
                    "Ghi chú": st.column_config.TextColumn("Ghi chú")
                },
                hide_index=True,
                key="cart_editor"
            )
            
            # --- XÁC NHẬN / HỦY ---
            col_x, col_y = st.columns([1, 4])
            with col_x:
                if st.button("✅ Xác nhận tất cả", type="primary"):
                    # Lặp qua dữ liệu đã chỉnh sửa từ bảng (edited_df)
                    for _, row in edited_df.iterrows():
                        t_controller.process_transaction(row["Mã"], row["Loại"], row["Số lượng"], str(row["Ghi chú"]) if pd.notna(row["Ghi chú"]) else "")
                    
                    st.session_state.cart = [] # Xóa giỏ hàng
                    clear_all_caches() # Chỉ clear cache khi lưu vào Google Sheet
                    st.success("🎉 Giao dịch thành công!")
                    st.rerun()
            with col_y:
                if st.button("🗑️ Hủy lưới"):
                    st.session_state.cart = []
                    st.rerun()

# --- TAB 3: BÁO CÁO TỒN KHO ---
elif menu == "Báo cáo tồn kho":
    show_report()

# --- TAB 4: LỊCH SỬ ---
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    history = get_cached_history(service)
    if history:
        st.dataframe(pd.DataFrame(history, columns=["Ngày", "Mã HH", "Loại", "Số Lượng", "Ghi Chú"]), width='stretch', hide_index=True)