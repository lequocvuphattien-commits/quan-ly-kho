import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report

# Cấu hình trang (Luôn để đầu tiên)
st.set_page_config(page_title="Quản Lý Kho", layout="wide")

# CSS tinh chỉnh màu sắc nút bấm và giao diện
st.markdown("""
    <style>
    /* Làm nổi bật nút Xác nhận tất cả (Xanh lá) */
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
    }
    /* Chỉnh sửa khoảng cách cho gọn gàng */
    .block-container { padding-top: 0rem !important; }
    </style>
""", unsafe_allow_html=True)

# Khởi tạo dịch vụ
service = DataService(mode="ONLINE")

# --- BỘ NHỚ ĐỆM (CACHE) ---
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc):
    return _svc.get_products()

@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc):
    return _svc.get_history()

st.title("📦 Quản lý kho")

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
                service.add_product(code.upper(), name, unit)
                st.cache_data.clear()
                st.success("Đã thêm thành công!"); st.rerun()

# --- TAB 2: NHẬP/XUẤT (PHIÊN BẢN TỐI ƯU GIAO DIỆN & TỐC ĐỘ) ---
elif menu == "Nhập/Xuất":
    st.header("🔄 Nhập/Xuất kho")
    
    # Khởi tạo giỏ hàng trong bộ nhớ tạm (Session State) để chống load chậm
    if 'cart' not in st.session_state: 
        st.session_state.cart = []
    
    products = get_cached_products(service)
    
    if products:
        # Chuẩn hóa dữ liệu thành dictionary để truy xuất siêu nhanh
        # p[1] là Mã, p[2] là Tên, p[3] là Đvt, p[4] là Tồn
        p_dict = {f"{p[1]} - {p[2]}": {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": p[4]} for p in products}
        
        # --- 1. KHUNG NHẬP LIỆU GIAO DỊCH ---
        with st.container(border=True):
            # Loại: Nút Button Nhập/Xuất (dạng radio ngang)
            trans_type = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
            
            # Chọn hàng: List box hỗ trợ gõ tìm kiếm tự động
            selected = st.selectbox(
                "Chọn hàng hóa", 
                options=list(p_dict.keys()), 
                index=None, 
                placeholder="🔍 Gõ tìm kiếm mã hoặc tên hàng..."
            )
            
            # Chia làm 3 cột: Số lượng, Ghi chú, Tồn
            c1, c2, c3 = st.columns([1, 1.5, 1])
            with c1:
                # Số lượng: Ô trống mặc định, tự động format có dấu phẩy
                qty = st.number_input("Số lượng", min_value=1.0, value=None, step=1.0, format="%.0f", placeholder="Nhập số...")
            with c2:
                # Ghi chú: Để trống nhập nếu cần
                note = st.text_input("Ghi chú", placeholder="Nhập ghi chú (tùy chọn)...")
            with c3:
                # Hiển thị Tồn động
                if selected:
                    current_stock = float(p_dict[selected]['Tồn'])
                    unit = p_dict[selected]['Đvt']
                    st.markdown(f"<div style='padding-top: 30px; font-size: 18px; font-weight: bold; color: #28a745;'>Tồn: {current_stock:,.0f} {unit}</div>", unsafe_allow_html=True)
                else:
                    st.markdown("<div style='padding-top: 30px; font-size: 18px;'>Tồn: --</div>", unsafe_allow_html=True)

            # Thêm vào lưới (TỐI ƯU HIỆU SUẤT X10)
            if st.button("➕ Thêm vào lưới"):
                if not selected or not qty:
                    st.warning("⚠️ Vui lòng chọn hàng hóa và nhập số lượng!")
                else:
                    prod_data = p_dict[selected]
                    
                    # Logic kiểm tra tồn bằng PYTHON THUẦN (siêu nhanh, không dùng Pandas)
                    if trans_type == "Xuất":
                        current_stock = float(prod_data["Tồn"])
                        
                        # Dùng generator expression của Python để tính tổng cực nhanh
                        out_in_cart = sum(
                            item["Số lượng"] 
                            for item in st.session_state.cart 
                            if item["Mã"] == prod_data["Mã"] and item["Loại"] == "Xuất"
                        )
                        
                        if qty + out_in_cart > current_stock:
                            st.error(f"❌ Không đủ tồn kho! (Tồn hiện tại: {current_stock:,.0f})")
                            st.stop()

                    # Đưa vào RAM
                    st.session_state.cart.append({
                        "Mã": prod_data["Mã"],
                        "Loại": trans_type,
                        "Đvt": prod_data["Đvt"],
                        "Số lượng": float(qty),
                        "Ghi chú": note if note else ""
                    })
                    st.rerun() # Load lại UI lập tức

        # --- 2. HIỂN THỊ LƯỚI CHỜ (CÓ THỂ SỬA TRỰC TIẾP) ---
        if st.session_state.cart:
            st.markdown("### 📋 Lưới chờ xử lý")
            st.caption("💡 *Mẹo: Click đúp chuột vào ô Số lượng hoặc Ghi chú để sửa trực tiếp trên bảng trước khi Xác nhận.*")
            
            # Data Editor thay thế dataframe thông thường
            edited_df = st.data_editor(
                pd.DataFrame(st.session_state.cart),
                column_config={
                    "Mã": st.column_config.TextColumn("Mã HH", disabled=True),
                    "Loại": st.column_config.TextColumn("Loại", disabled=True),
                    "Đvt": st.column_config.TextColumn("Đvt", disabled=True),
                    "Số lượng": st.column_config.NumberColumn("Số lượng", required=True, min_value=1.0, format="%.0f"),
                    "Ghi chú": st.column_config.TextColumn("Ghi chú")
                },
                hide_index=True,
                use_container_width=True,
                key="cart_editor"
            )
            
            # --- 3. XÁC NHẬN VÀ HỦY ---
            col_x, col_y = st.columns([1, 5])
            with col_x:
                if st.button("✅ Xác nhận tất cả", type="primary"): # type="primary" sẽ kích hoạt CSS màu xanh lá
                    for _, row in edited_df.iterrows():
                        service.add_transaction(row["Mã"], row["Số lượng"], row["Loại"], str(row["Ghi chú"]) if pd.notna(row["Ghi chú"]) else "")
                        service.update_stock(row["Mã"], row["Số lượng"], row["Loại"])
                    
                    st.session_state.cart = []
                    st.cache_data.clear()
                    st.success("🎉 Đã lưu toàn bộ giao dịch vào hệ thống!")
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