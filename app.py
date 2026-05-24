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
        # --- BẮT ĐẦU ĐOẠN FIX LỖI TYPE ERROR ---
        # Code tự động nhận diện kiểu dữ liệu của 'p' để xử lý an toàn
        p_dict = {}
        for p in products:
            if isinstance(p, dict):
                # Nếu p đã là dictionary
                p_dict[f"{p.get('Mã', '')} - {p.get('Tên', '')}"] = p
            elif hasattr(p, 'code'):
                # Nếu p là đối tượng (Object) từ ProductController
                p_dict[f"{p.code} - {p.name}"] = {"Mã": p.code, "Tên": p.name, "Đvt": p.unit, "Tồn": p.stock}
            else:
                # Nếu p là list hoặc tuple (ví dụ: [ID, Mã, Tên, Đvt, Tồn])
                p_dict[f"{p[1]} - {p[2]}"] = {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": p[4]}
        # --- KẾT THÚC ĐOẠN FIX LỖI ---
        
        # --- KHUNG NHẬP LIỆU GIAO DỊCH ---
        with st.container(border=True):
            typ = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
            
            sel = st.selectbox(
                "Chọn hàng hóa", 
                options=list(p_dict.keys()), 
                index=None, 
                placeholder="🔍 Gõ tìm kiếm mã hoặc tên hàng..."
            )
            
            c1, c2, c3 = st.columns([1, 1.5, 1])
            with c1:
                qty = st.number_input("Số lượng", min_value=1, value=None, step=1, format="%d", placeholder="Nhập số...")
            with c2:
                note = st.text_input("Ghi chú", placeholder="Nhập ghi chú...")
            with c3:
                if sel:
                    current_stock = float(p_dict[sel]['Tồn'])
                    unit = p_dict[sel]['Đvt']
                    st.markdown(f"<div style='padding-top: 30px; font-size: 18px; font-weight: bold; color: #28a745;'>Tồn: {current_stock:,.0f} {unit}</div>", unsafe_allow_html=True)
                else:
                    st.markdown("<div style='padding-top: 30px; font-size: 18px;'>Tồn: --</div>", unsafe_allow_html=True)

            if st.button("➕ Thêm vào lưới"):
                if not sel or not qty:
                    st.warning("⚠️ Vui lòng chọn hàng hóa và nhập số lượng!")
                else:
                    prod_data = p_dict[sel]
                    prod_code = prod_data["Mã"]
                    current_stock = float(prod_data["Tồn"])

                    if typ == "Xuất":
                        cart_df = pd.DataFrame(st.session_state.cart) if st.session_state.cart else pd.DataFrame(columns=["Mã", "Loại", "Số lượng"])
                        out_in_cart = cart_df[(cart_df["Mã"] == prod_code) & (cart_df["Loại"] == "Xuất")]["Số lượng"].sum() if not cart_df.empty else 0
                        if qty + out_in_cart > current_stock:
                            st.error(f"❌ Không đủ tồn kho! (Tồn hiện tại: {current_stock:,.0f})")
                            st.stop()

                    st.session_state.cart.append({
                        "Mã": prod_code,
                        "Tên HH": prod_data["Tên"],
                        "Đvt": prod_data["Đvt"],
                        "Loại": typ,
                        "Số lượng": qty,
                        "Ghi chú": note if note else ""
                    })
                    st.rerun()

        # --- HIỂN THỊ LƯỚI CHỜ ---
        if st.session_state.cart:
            st.markdown("### 📋 Lưới chờ xử lý")
            st.caption("💡 *Mẹo: Click đúp chuột vào ô Số lượng hoặc Ghi chú để sửa trực tiếp trên bảng.*")
            
            edited_df = st.data_editor(
                pd.DataFrame(st.session_state.cart),
                column_config={
                    "Mã": st.column_config.TextColumn("Mã HH", disabled=True),
                    "Tên HH": st.column_config.TextColumn("Tên HH", disabled=True),
                    "Đvt": st.column_config.TextColumn("Đvt", disabled=True),
                    "Loại": st.column_config.TextColumn("Loại", disabled=True),
                    "Số lượng": st.column_config.NumberColumn("Số lượng", required=True, min_value=1, format="%d"),
                    "Ghi chú": st.column_config.TextColumn("Ghi chú")
                },
                hide_index=True,
                key="cart_editor"
            )
            
            col_x, col_y = st.columns([1, 4])
            with col_x:
                if st.button("✅ Xác nhận tất cả", type="primary"):
                    for _, row in edited_df.iterrows():
                        t_controller.process_transaction(row["Mã"], row["Loại"], row["Số lượng"], str(row["Ghi chú"]) if pd.notna(row["Ghi chú"]) else "")
                    
                    st.session_state.cart = []
                    clear_all_caches()
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