import streamlit as st
import pandas as pd
from services.data_service import DataService
from views.report_view_streamlit import show_report

# Khởi tạo dịch vụ
service = DataService(mode="ONLINE")

# --- BỘ NHỚ ĐỆM (CACHE) TỐI ƯU TỐC ĐỘ ---
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc):
    return _svc.get_products()

@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc):
    return _svc.get_history()

# Cấu hình trang
st.set_page_config(page_title="Quản Lý Kho Hàng", layout="wide")
st.title("📦 Quản lý kho hàng")

# Menu điều hướng
menu = st.sidebar.selectbox("Menu", ["Danh mục hàng hóa", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 1: DANH MỤC HÀNG HÓA ---
if menu == "Danh mục hàng hóa":
    st.header("Danh mục hàng hóa")
    products = get_cached_products(service)
    
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        
        # Hiển thị bảng
        st.dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]], use_container_width=True, hide_index=True)
        
        # Nút xuất Excel (CSV)
        csv = df[["Mã", "Tên", "Đvt", "Tồn"]].to_csv(index=False).encode('utf-8')
        st.download_button(label="📥 Xuất danh mục ra Excel (CSV)", data=csv, file_name="DanhMucHangHoa.csv", mime="text/csv")
    
    with st.form("add_form", clear_on_submit=True):
        st.subheader("Thêm hàng hóa mới")
        code, name, unit = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
        if st.form_submit_button("Thêm hàng hóa"):
            if not code or not name: 
                st.warning("Nhập đủ Mã và Tên!")
            elif service.check_product_exists(code.upper()): 
                st.error("Mã đã tồn tại!")
            else:
                service.add_product(code, name, unit)
                st.cache_data.clear()
                st.success("Đã thêm thành công!"); st.rerun()

# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất":
    st.header("Nhập/Xuất kho")
    products = get_cached_products(service)
    
    if products:
        product_dict = {f"{p[1]} - {p[2]}": p[1] for p in products}
        selected = st.selectbox("Chọn hàng", list(product_dict.keys()))
        prod_code = product_dict[selected]
        qty = st.number_input("Số lượng", min_value=0.0, step=1.0)
        trans_type = st.radio("Loại", ["Nhập", "Xuất"])
        note = st.text_input("Ghi chú")
        
        if st.button("Xác nhận giao dịch"):
            service.add_transaction(prod_code, qty, trans_type, note)
            service.update_stock(prod_code, qty, trans_type)
            st.cache_data.clear()
            st.success("Đã cập nhật tồn kho!"); st.rerun()

# --- TAB 3: BÁO CÁO TỒN KHO ---
elif menu == "Báo cáo tồn kho":
    show_report()

# --- TAB 4: LỊCH SỬ ---
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    history = get_cached_history(service)
    if history:
        st.dataframe(pd.DataFrame(history, columns=["Ngày", "Mã HH", "Loại", "Số Lượng", "Ghi Chú"]), use_container_width=True)