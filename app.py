import streamlit as st
import pandas as pd
from services.data_service import DataService
from controllers.product_controller import ProductController
from controllers.transaction_controller import TransactionController
from utils.export import export_to_excel

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Quản Lý Kho", layout="wide")

# CSS tối ưu giao diện
st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; }
    .blue-btn button { background-color: #007BFF !important; color: white !important; width: 100%; }
    .green-btn button { background-color: #28a745 !important; color: white !important; width: 100%; }
    </style>
""", unsafe_allow_html=True)

# Khởi tạo các Controller
service = DataService()
p_controller = ProductController()
t_controller = TransactionController()

def clear_all():
    st.cache_data.clear()

# CACHE DỮ LIỆU ĐÃ CHUYỂN ĐỔI SANG DICT
@st.cache_data(ttl=60)
def get_cached_products(_service):
    # Chuyển đổi Object từ Controller sang list of dicts để Streamlit lưu cache được
    products = p_controller.get_all_products()
    return [vars(p) if hasattr(p, '__dict__') else p for p in products]

st.title("📦 Hệ Thống Quản Lý Kho")
menu = st.sidebar.radio("Menu", ["Danh mục HH", "Báo cáo tồn kho", "Giao dịch"])

# --- TAB 1: DANH MỤC HH ---
if menu == "Danh mục HH":
    st.header("Danh mục hàng hóa")
    products = get_cached_products(service)
    
    if products:
        df = pd.DataFrame(products)
        # Chọn các cột hiển thị: code, name, unit, stock
        cols_to_show = ["code", "name", "unit", "stock"] if "code" in df.columns else df.columns
        st.dataframe(df[cols_to_show], use_container_width=True, hide_index=True)
        
        st.download_button("📥 Xuất Excel", export_to_excel(df), "DanhMuc.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Danh mục trống!")

    st.divider()
    with st.form("add_product_form", clear_on_submit=True):
        st.subheader("➕ Thêm hàng mới")
        c1, c2, c3 = st.columns(3)
        code = c1.text_input("Mã hàng")
        name = c2.text_input("Tên hàng")
        unit = c3.text_input("ĐVT")
        
        if st.form_submit_button("Lưu hàng hóa"):
            if code and name:
                p_controller.add_product(code.upper(), name, unit)
                clear_all()
                st.success("Thêm thành công!")
                st.rerun()
            else:
                st.error("Nhập thiếu thông tin!")

# --- TAB 2: BÁO CÁO TỒN KHO ---
elif menu == "Báo cáo tồn kho":
    st.header("Báo cáo tồn kho")
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Từ ngày")
    end_date = col2.date_input("Đến ngày")
    
    if st.button("Lọc báo cáo"):
        st.info("Đang xử lý dữ liệu...")
        # Ở đây bạn gọi hàm xử lý báo cáo từ Controller
        # t_controller.get_product_stats_by_date(...)

# --- TAB 3: GIAO DỊCH ---
elif menu == "Giao dịch":
    st.header("Ghi nhận giao dịch")
    products = get_cached_products(service)
    prod_list = [p['code'] for p in products]
    
    with st.form("trans_form"):
        code = st.selectbox("Chọn mã hàng", prod_list)
        type_ = st.selectbox("Loại", ["Nhập", "Xuất"])
        qty = st.number_input("Số lượng", min_value=1)
        note = st.text_input("Ghi chú")
        
        if st.form_submit_button("Xác nhận"):
            res = t_controller.process_transaction(code, type_, qty, note)
            if res == True:
                st.success("Giao dịch thành công!")
                clear_all()
                st.rerun()
            elif res == "ERROR_INSUFFICIENT_STOCK":
                st.error("Không đủ tồn kho để xuất!")
            else:
                st.error("Có lỗi xảy ra!")