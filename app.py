import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

# --- BỘ NHỚ ĐỆM ---
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_products(_svc): return _svc.get_products()
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_history(_svc): return _svc.get_history()
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_config(_svc): return _svc.get_config_options()
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_employees(_svc): return _svc.get_employees()

st.set_page_config(page_title="Quản Lý Kho", layout="wide", initial_sidebar_state="collapsed")

# --- CSS TỐI ƯU GIAO DIỆN ---
st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    div.stButton > button[kind="primary"] { background-color: #28a745 !important; color: white !important; }
    h1 { padding-bottom: 0rem !important; margin-bottom: 0rem !important; }
    h3 { padding-top: 0rem !important; margin-top: 0rem !important; }
    div[data-testid="stSelectbox"] { margin-bottom: -1rem !important; }
    </style>
""", unsafe_allow_html=True)

@st.cache_resource(show_spinner=False)
def get_data_service(): return DataService(mode="ONLINE")

service = get_data_service()
st.title("📦 Quản lý kho")

# --- ĐĂNG NHẬP ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if not st.session_state.logged_in:
    with st.container(border=True):
        st.subheader("🔒 Đăng nhập hệ thống")
        pwd = st.text_input("Mật khẩu:", type="password") 
        if st.button("Đăng nhập", type="primary"):
            if pwd == "123": st.session_state.logged_in = True; st.rerun() 
            else: st.error("❌ Mật khẩu không đúng!")
    st.stop() 

menu = st.selectbox("Chức năng", ["Danh mục hàng", "Nhập/Xuất Kho", "Báo cáo tồn kho", "Lịch sử giao dịch", "Quản lý nhân viên"], label_visibility="collapsed")

# --- TAB DANH MỤC HÀNG ---
if menu == "Danh mục hàng":
    st.subheader("📋 Danh mục hàng")
    products = get_cached_products(service)
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên hàng hóa", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        gb = GridOptionsBuilder.from_dataframe(df[["Mã", "Tên hàng hóa", "Đvt", "Tồn"]])
        gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1)
        gb.configure_column("Mã", minWidth=50, editable=False)
        gb.configure_column("Tên hàng hóa", minWidth=150, editable=True, cellStyle={'backgroundColor': '#f0f8ff'})
        gb.configure_column("Đvt", minWidth=50, editable=True, cellStyle={'backgroundColor': '#f0f8ff'})
        gb.configure_column("Tồn", minWidth=60, editable=False, type=["numericColumn"], valueFormatter="Number(x).toLocaleString('en-US')")
        
        grid_response = AgGrid(df[["Mã", "Tên hàng hóa", "Đvt", "Tồn"]], gridOptions=gb.build(), fit_columns_on_grid_load=True, theme='streamlit', update_mode=GridUpdateMode.MODEL_CHANGED, height=400)
        
        # (Lưu thay đổi ở đây giống code cũ của bạn)

# --- TAB NHẬP/XUẤT ---
elif menu == "Nhập/Xuất Kho":
    st.subheader("🔄 Nhập/Xuất kho")
    trans_type = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
    kho_nhap_list, kho_xuat_list = service.get_config_options()
    products = get_cached_products(service)
    employees = get_cached_employees(service)
    
    if products:
        p_dict = {f"{p[1]} - {p[2]} (Tồn: {float(p[4]):,.0f} {p[3]})": {"Mã": p[1], "Tên hàng hóa": p[2], "Đvt": p[3], "Tồn": p[4]} for p in products}
        
        with st.container(border=True):
            selected = st.selectbox("Chọn hàng hóa", options=list(p_dict.keys()), index=None)
            c1, c2, c3 = st.columns([1.5, 2, 1])
            with c1: qty = st.number_input("Số lượng", min_value=1.0, value=None, step=1.0)
            with c2: note = st.selectbox("Diễn giải / Kho", options=(kho_nhap_list if trans_type == "Nhập" else kho_xuat_list), index=None)
            with c3:
                if st.button("➕ Thêm vào lưới"):
                    if not selected or not qty or not note: st.warning("⚠️ Nhập đủ thông tin!")
                    else:
                        st.session_state.cart.append({"Mã HH": p_dict[selected]["Mã"], "Tên HH": p_dict[selected]["Tên hàng hóa"], "Đvt": p_dict[selected]["Đvt"], "Số lượng": float(qty), "Ghi chú": note, "Loại": trans_type})
                        st.rerun()

        if 'cart' not in st.session_state: st.session_state.cart = []
        if st.session_state.cart:
            edited_df_cart = st.data_editor(pd.DataFrame(st.session_state.cart), use_container_width=True)
            
            # --- CHỌN NHÂN VIÊN TRƯỚC KHI XÁC NHẬN ---
            emp_list = [emp[1] for emp in employees]
            selected_emp = st.selectbox("Nhân viên thực hiện:", options=emp_list, index=None, placeholder="Chọn tên bạn...")
            
            if st.button("✅ Xác nhận tất cả", type="primary"): 
                if not selected_emp: st.warning("⚠️ Vui lòng chọn nhân viên thực hiện!")
                else:
                    for _, row in edited_df_cart.iterrows():
                        service.add_transaction(row["Mã HH"], row["Tên HH"], row["Số lượng"], row["Loại"], row["Ghi chú"], selected_emp)
                        service.update_stock(row["Mã HH"], row["Số lượng"], row["Loại"])
                    st.session_state.cart = []
                    st.cache_data.clear()
                    st.success(f"🎉 Giao dịch thành công bởi {selected_emp}!")
                    st.rerun()

elif menu == "Báo cáo tồn kho": show_report()

# --- TAB LỊCH SỬ GIAO DỊCH ---
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    history = get_cached_history(service)
    if history:
        df = pd.DataFrame(history, columns=["Ngày", "Mã", "Tên hàng hóa", "Loại", "Số Lượng", "Ghi Chú", "Nhân viên"])
        AgGrid(df, fit_columns_on_grid_load=True, theme='streamlit', height=650)

# =====================================================================
# --- [THÊM MỚI]: TAB 5: QUẢN LÝ NHÂN VIÊN ---
# =====================================================================
elif menu == "Quản lý nhân viên":
    st.subheader("👥 Quản lý nhân viên")
    employees = get_cached_employees(service)
    
    if employees:
        df_emp = pd.DataFrame(employees, columns=["Mã NV", "Tên NV", "Số điện thoại", "Chức vụ"])
        
        # Cấu hình AgGrid cho phép Sửa trực tiếp (Trừ Mã NV)
        gb = GridOptionsBuilder.from_dataframe(df_emp)
        gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1)
        
        gb.configure_column("Mã NV", minWidth=80, editable=False, cellStyle={'textAlign': 'center'})
        gb.configure_column("Tên NV", minWidth=150, editable=True, cellStyle={'textAlign': 'left', 'backgroundColor': '#f0f8ff'}) 
        gb.configure_column("Số điện thoại", minWidth=120, editable=True, cellStyle={'textAlign': 'center', 'backgroundColor': '#f0f8ff'})
        gb.configure_column("Chức vụ", minWidth=120, editable=True, cellStyle={'textAlign': 'center', 'backgroundColor': '#f0f8ff'})
        
        go = gb.build()
        
        grid_response = AgGrid(
            df_emp,
            gridOptions=go,
            fit_columns_on_grid_load=True,
            theme='streamlit',
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            height=400
        )

        edited_df_emp = pd.DataFrame(grid_response['data'])

        # --- XỬ LÝ LƯU THAY ĐỔI NHÂN VIÊN TỪ AGGRID ---
        has_changes = False
        changes_to_save = [] 
        
        if not edited_df_emp.empty:
            for i in range(len(edited_df_emp)):
                ma = edited_df_emp.iloc[i]["Mã NV"]
                ten_moi = edited_df_emp.iloc[i]["Tên NV"]
                sdt_moi = edited_df_emp.iloc[i]["Số điện thoại"]
                cv_moi = edited_df_emp.iloc[i]["Chức vụ"]
                
                orig_row = df_emp[df_emp["Mã NV"] == ma].iloc[0]
                
                if ten_moi != orig_row["Tên NV"] or sdt_moi != orig_row["Số điện thoại"] or cv_moi != orig_row["Chức vụ"]:
                    has_changes = True
                    changes_to_save.append({"Mã NV": ma, "Tên NV": ten_moi, "Số điện thoại": sdt_moi, "Chức vụ": cv_moi})

        if has_changes:
            st.info("⚠️ Có thay đổi chưa được lưu!")
            if st.button("💾 Lưu thay đổi", type="primary"):
                for item in changes_to_save:
                    service.update_employee(item["Mã NV"], item["Tên NV"], item["Số điện thoại"], item["Chức vụ"])
                
                st.cache_data.clear()
                st.success("🎉 Đã cập nhật thông tin nhân viên thành công!")
                st.rerun()

    # --- GIAO DIỆN THÊM VÀ XÓA NHÂN VIÊN NẰM CẠNH NHAU ---
    c1, c2 = st.columns(2)
    with c1:
        with st.expander("➕ Thêm nhân viên mới"):
            with st.form("add_emp_form", clear_on_submit=True):
                emp_code = st.text_input("Mã nhân viên (VD: NV01)")
                emp_name = st.text_input("Tên nhân viên")
                emp_phone = st.text_input("Số điện thoại")
                emp_role = st.selectbox("Chức vụ", ["Nhân viên kho", "Quản lý", "Kế toán", "Tài xế", "Khác"])
                
                if st.form_submit_button("Thêm nhân viên"):
                    if not emp_code or not emp_name: 
                        st.warning("Vui lòng nhập đủ Mã và Tên nhân viên!")
                    elif service.check_employee_exists(emp_code.upper()): 
                        st.error("Mã nhân viên này đã tồn tại!")
                    else:
                        service.add_employee(emp_code.upper(), emp_name, emp_phone, emp_role)
                        st.cache_data.clear()
                        st.success("Đã thêm nhân viên thành công!")
                        st.rerun()
    with c2:
        with st.expander("🗑️ Xóa nhân viên"):
            if employees:
                del_emp_code = st.selectbox("Chọn mã NV cần xóa", options=df_emp["Mã NV"].tolist())
                if st.button("Xác nhận xóa nhân viên"):
                    service.delete_employee(del_emp_code)
                    st.cache_data.clear()
                    st.success(f"Đã xóa nhân viên {del_emp_code}!")
                    st.rerun()