import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode

# --- BỘ NHỚ ĐỆM (CACHE) TỐI ƯU TỐC ĐỘ ---
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_products(_svc):
    return _svc.get_products()

@st.cache_data(ttl=600, show_spinner=False)
def get_cached_history(_svc):
    return _svc.get_history()

# Thêm bộ nhớ đệm cho cấu hình Kho (Tối ưu độ trễ)
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_config(_svc):
    return _svc.get_config_options()

# Cấu hình trang
st.set_page_config(page_title="Quản Lý Kho", layout="wide")

st.markdown("""
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;}        
    </style>
""", unsafe_allow_html=True)

service = DataService(mode="ONLINE")
st.title("📦 Quản lý kho")
menu = st.sidebar.selectbox("Menu", ["Danh mục hàng", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 1: DANH MỤC HÀNG ---
if menu == "Danh mục hàng":
    st.subheader("📋 Danh mục hàng")
    products = get_cached_products(service)
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        gb = GridOptionsBuilder.from_dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]])
        gb.configure_default_column(sortable=True, filter=True, floatingFilter=True, resizable=True, suppressMenu=True)
        gb.configure_pagination(paginationAutoPageSize=True)
        go = gb.build()
        AgGrid(df[["Mã", "Tên", "Đvt", "Tồn"]], gridOptions=go, fit_columns_on_grid_load=True, theme='streamlit')
    
    with st.expander("➕ Thêm hàng hóa mới"):
        with st.form("add_form", clear_on_submit=True):
            code, name, unit = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
            if st.form_submit_button("Thêm hàng hóa"):
                if not code or not name: st.warning("Nhập đủ Mã và Tên!")
                elif service.check_product_exists(code.upper()): st.error("Mã đã tồn tại!")
                else:
                    service.add_product(code.upper(), name, unit)
                    st.cache_data.clear(); st.success("Đã thêm thành công!"); st.rerun()

# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất":
    st.subheader("🔄 Nhập/Xuất kho")
    trans_type = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
    
    # SỬ DỤNG CACHE ĐỂ LẤY DANH SÁCH KHO (Tăng tốc độ gấp 10 lần)
    kho_nhap_list, kho_xuat_list = get_cached_config(service)
    
    with st.expander(f"➕ Thêm địa điểm mới vào {trans_type}"):
        new_kho = st.text_input("Tên địa điểm mới", key="new_kho_input")
        if st.button("Lưu địa điểm mới"):
            if new_kho:
                service.add_config_option(trans_type, new_kho)
                st.cache_data.clear(); st.rerun() # Xóa cache để tải lại danh sách mới
            else: st.warning("Vui lòng nhập tên!")

    products = get_cached_products(service)
    if products:
        p_dict = {f"{p[1]} - {p[2]} (Tồn: {float(p[4]):,.0f} {p[3]})": {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": p[4]} for p in products}
        
        with st.form("input_form", clear_on_submit=True):
            selected = st.selectbox("Chọn hàng hóa", options=list(p_dict.keys()), index=None)
            c1, c2 = st.columns([1, 1])
            qty = c1.number_input("Số lượng", min_value=1.0, step=1.0, format="%.0f")
            note = c2.selectbox("Diễn giải / Kho", options=(kho_nhap_list if trans_type == "Nhập" else kho_xuat_list), index=None)
            submitted = st.form_submit_button("➕ Thêm vào lưới")
        
        if submitted:
            if not selected or not note: st.warning("Chọn đủ thông tin!"); st.stop()
            prod = p_dict[selected]
            if trans_type == "Xuất" and qty > float(prod["Tồn"]): st.error("Không đủ tồn!"); st.stop()
            if 'cart' not in st.session_state: st.session_state.cart = []
            st.session_state.cart.append({"Mã HH": prod["Mã"], "Tên HH": prod["Tên"], "Đvt": prod["Đvt"], "Số lượng": float(qty), "Ghi chú": note, "Loại": trans_type})
            st.rerun()

    if 'cart' in st.session_state and st.session_state.cart:
        edited_df = st.data_editor(pd.DataFrame(st.session_state.cart), hide_index=True, use_container_width=True)
        if st.button("✅ Xác nhận tất cả", type="primary"):
            for _, row in edited_df.iterrows():
                service.add_transaction(row["Mã HH"], row["Tên HH"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                service.update_stock(row["Mã HH"], row["Số lượng"], row["Loại"])
            st.session_state.cart = []; st.cache_data.clear(); st.success("Thành công!"); st.rerun()

elif menu == "Báo cáo tồn kho": show_report()

elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    history = get_cached_history(service)
    if history:
        df = pd.DataFrame(history, columns=["Ngày", "Mã", "Tên Hàng Hóa", "Loại", "Số Lượng", "Ghi Chú"])
        df["Số Lượng"] = pd.to_numeric(df["Số Lượng"], errors="coerce").fillna(0)
        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_column("Số Lượng", type=["numericColumn"], headerClass='ag-right-aligned-header', cellClass='ag-right-aligned-cell')
        gb.configure_column("Ghi Chú", headerClass='ag-right-aligned-header', cellStyle={'textAlign': 'right'})
        AgGrid(df, gridOptions=gb.build(), fit_columns_on_grid_load=True, theme='streamlit')