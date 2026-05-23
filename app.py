import streamlit as st
import pandas as pd
from services.data_service import DataService
from controllers.product_controller import ProductController
from controllers.transaction_controller import TransactionController
from utils.export import export_to_excel
from views.report_view_streamlit import show_report

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Quản Lý Kho", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div.stButton > button { width: 100%; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Quản lý kho hàng")

# 2. KHỞI TẠO DỊCH VỤ
service = DataService(mode="ONLINE")
p_controller = ProductController()
t_controller = TransactionController()

# 3. HÀM CACHE DỮ LIỆU ĐÃ FIX LỖI SERIEALIZATION
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc):
    # Lấy danh sách đối tượng
    products = p_controller.get_all_products()
    
    # CHUYỂN ĐỔI AN TOÀN: Chỉ lấy dữ liệu thô (string/float/int)
    # Không dùng vars() hay __dict__ để tránh lỗi descriptor
    clean_data = []
    for p in products:
        item = {
            "ID": getattr(p, 'id', 'N/A'),
            "Mã": getattr(p, 'code', 'N/A'),
            "Tên": getattr(p, 'name', 'N/A'),
            "Đvt": getattr(p, 'unit', 'N/A'),
            "Tồn": float(getattr(p, 'stock', 0))
        }
        clean_data.append(item)
    return clean_data

@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc):
    return t_controller.get_transaction_history()

def clear_all_caches():
    st.cache_data.clear()

# 4. MENU ĐIỀU HƯỚNG
menu = st.sidebar.selectbox("Menu", ["Danh mục hàng hóa", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# 5. CÁC TAB CHỨC NĂNG
if menu == "Danh mục hàng hóa":
    st.header("📋 Danh mục hàng hóa")
    products = get_cached_products(service)
    if products:
        df = pd.DataFrame(products)
        # Sử dụng width='stretch' để tránh cảnh báo
        st.dataframe(df, width=None, hide_index=True)
        st.download_button("📥 Xuất danh mục (.xlsx)", export_to_excel(df), "DanhMuc.xlsx")
    
    with st.expander("➕ Thêm hàng hóa mới"):
        with st.form("add_form", clear_on_submit=True):
            c, n, u = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
            if st.form_submit_button("Lưu hàng hóa"):
                if not c or not n: st.warning("Nhập đủ Mã và Tên!")
                else:
                    success, msg = p_controller.add_product(c.upper(), n, u)
                    if success:
                        st.success(msg)
                        clear_all_caches()
                        st.rerun()
                    else: st.error(msg)

elif menu == "Nhập/Xuất":
    st.header("🔄 Nhập/Xuất hàng loạt")
    if 'cart' not in st.session_state: st.session_state.cart = []
    
    products = get_cached_products(service)
    if products:
        # p['Mã'] là key, p['Tồn'] là value
        stock_map = {str(p['Mã']).strip(): float(p['Tồn']) for p in products}
        p_dict = {f"{p['Mã']} - {p['Tên']}": p['Mã'] for p in products}
        
        col1, col2, col3 = st.columns([2, 1, 1])
        sel = col1.selectbox("Chọn hàng", list(p_dict.keys()))
        qty = col2.number_input("Số lượng", min_value=1.0, step=1.0)
        typ = col3.radio("Loại", ["Nhập", "Xuất"], horizontal=True)
        
        if st.button("➕ Thêm vào lưới"):
            prod_code = p_dict[sel]
            if typ == "Xuất":
                current_stock = stock_map.get(str(prod_code).strip(), 0)
                cart_df = pd.DataFrame(st.session_state.cart) if st.session_state.cart else pd.DataFrame(columns=["Mã", "Loại", "Số lượng"])
                out_in_cart = cart_df[(cart_df["Mã"] == prod_code) & (cart_df["Loại"] == "Xuất")]["Số lượng"].sum()
                if qty + out_in_cart > current_stock:
                    st.error(f"❌ Không đủ tồn kho! (Tồn: {current_stock})")
                else:
                    st.session_state.cart.append({"Mã": prod_code, "Loại": typ, "Số lượng": qty})
                    st.rerun()
            else:
                st.session_state.cart.append({"Mã": prod_code, "Loại": typ, "Số lượng": qty})
                st.rerun()

        if st.session_state.cart:
            st.write("📋 Danh sách chờ:")
            st.table(pd.DataFrame(st.session_state.cart))
            c1, c2 = st.columns(2)
            if c1.button("✅ Xác nhận tất cả"):
                for item in st.session_state.cart:
                    t_controller.process_transaction(item["Mã"], item["Loại"], item["Số lượng"], "Hàng loạt")
                st.session_state.cart = []
                clear_all_caches(); st.success("Hoàn tất!"); st.rerun()
            if c2.button("🗑️ Hủy lưới"):
                st.session_state.cart = []; st.rerun()

elif menu == "Báo cáo tồn kho":
    show_report()

elif menu == "Lịch sử giao dịch":
    st.header("📜 Lịch sử giao dịch")
    hist = get_cached_history(service)
    if hist:
        df_h = pd.DataFrame(hist, columns=["Ngày", "Mã", "Loại", "SL", "Ghi chú"])
        st.dataframe(df_h, width=None)
        st.download_button("📥 Xuất lịch sử Excel", export_to_excel(df_h), "LichSu.xlsx")
    else:
        st.info("Chưa có lịch sử giao dịch.")