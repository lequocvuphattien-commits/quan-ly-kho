import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report

# Cấu hình trang (Phải đặt đầu tiên)
st.set_page_config(page_title="Quản Lý Kho Hàng", layout="wide")

# Khởi tạo dịch vụ
service = DataService(mode="ONLINE")

# --- BỘ NHỚ ĐỆM (CACHE) TỐI ƯU ---
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc):
    return _svc.get_products()

@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc):
    return _svc.get_history()

def clear_all_caches():
    st.cache_data.clear()

# --- HÀM XUẤT EXCEL CHUYÊN NGHIỆP ---
def export_to_excel(df, sheet_name='Data'):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        worksheet.freeze_panes = 'A2'
        max_row, max_col = worksheet.max_row, worksheet.max_column
        worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
        for col in range(1, max_col + 1):
            worksheet.column_dimensions[get_column_letter(col)].width = 15
    return buffer.getvalue()

# --- GIAO DIỆN CHÍNH ---
st.title("📦 Quản lý kho hàng")
menu = st.sidebar.selectbox("Menu", ["Danh mục hàng hóa", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 1: DANH MỤC ---
if menu == "Danh mục hàng hóa":
    st.header("Danh mục hàng hóa")
    products = get_cached_products(service)
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        st.dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]], use_container_width=True, hide_index=True)
        st.download_button("📥 Xuất danh mục (.xlsx)", export_to_excel(df), "DanhMuc.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    with st.form("add_form", clear_on_submit=True):
        st.subheader("Thêm hàng hóa mới")
        c, n, u = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
        if st.form_submit_button("Thêm hàng hóa"):
            if not c or not n: st.warning("Nhập đủ Mã và Tên!")
            elif service.check_product_exists(c.upper()): st.error("Mã đã tồn tại!")
            else:
                service.add_product(c, n, u)
                clear_all_caches(); st.success("Đã thêm!"); st.rerun()

# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất":
    st.header("Nhập/Xuất hàng loạt")
    if 'cart' not in st.session_state: st.session_state.cart = []
    
    products = get_cached_products(service)
    if products:
        p_dict = {f"{p[1]} - {p[2]}": p[1] for p in products}
        col1, col2, col3 = st.columns([2, 1, 1])
        sel = col1.selectbox("Chọn hàng", list(p_dict.keys()))
        qty = col2.number_input("Số lượng", min_value=1.0, step=1.0)
        typ = col3.radio("Loại", ["Nhập", "Xuất"], horizontal=True)
        
        if st.button("➕ Thêm vào lưới"):
            st.session_state.cart.append({"Mã": p_dict[sel], "Loại": typ, "Số lượng": qty})
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

# --- TAB 3: BÁO CÁO ---
elif menu == "Báo cáo tồn kho":
    show_report()

# --- TAB 4: LỊCH SỬ ---
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    hist = get_cached_history(service)
    if hist:
        df_h = pd.DataFrame(hist, columns=["Ngày", "Mã", "Loại", "SL", "Ghi chú"])
        st.dataframe(df_h, use_container_width=True)
        st.download_button("📥 Xuất lịch sử (.xlsx)", export_to_excel(df_h), "LichSu.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")