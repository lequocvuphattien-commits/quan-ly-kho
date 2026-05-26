import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.print_export_view import show_print_export_view
from views.report_view_streamlit import show_report
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
import datetime

# --- BỘ NHỚ ĐỆM ---
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_products(_svc): return _svc.get_products()
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_history(_svc): return _svc.get_history()
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_config(_svc): return _svc.get_config_options()
@st.cache_data(ttl=600, show_spinner=False)
def get_cached_employees(_svc): return _svc.get_employees()

# --- ĐẶT HÀM DIALOG Ở ĐÂY (TRƯỚC KHI KHỞI TẠO SERVICE) ---
@st.dialog("Xác nhận xóa")
def confirm_delete(product_name, del_code, service):
    st.warning(f"Bạn có chắc muốn xóa: {product_name}?")
    col_yes, col_no = st.columns(2)
    with col_yes:
        if st.button("✅ Yes"):
            service.delete_product(del_code)
            st.cache_data.clear()
            st.success("Đã xóa!")
            st.rerun() 
    with col_no:
        if st.button("❌ No"):
            st.rerun()

st.set_page_config(page_title="Quản Lý Kho", layout="wide", initial_sidebar_state="collapsed")

@st.dialog("Xác nhận xóa lịch sử")
def confirm_delete_history_dialog(selected_rows, service):
    st.warning(f"⚠️ Bạn có chắc muốn xóa **{len(selected_rows)}** giao dịch đã chọn không?")
    st.info("Hành động này sẽ tự động HOÀN TÁC (cộng/trừ ngược) số lượng tồn kho tương ứng.")
    
    msg_container = st.container()
    
    col_yes, col_no = st.columns(2)
    with col_yes:
        if st.button("✅ Xác nhận xóa"):
            
            # ==========================================
            # BƯỚC 1: KIỂM TRA RÀNG BUỘC (SIÊU CHUẨN XÁC)
            # ==========================================
            # Lấy dữ liệu trực tiếp từ DB, BỎ QUA CACHE để có tồn kho mới nhất
            products = service.get_products()
            
            # Ép kiểu chuỗi và loại bỏ khoảng trắng ở 2 đầu để so sánh khớp 100%
            stock_dict = {str(p[1]).strip(): float(p[4]) for p in products} if products else {}
            
            stock_changes = {}
            for index, row in selected_rows.iterrows():
                # Tự động nhận diện tên cột là "Mã HH" hay "Mã"
                p_code = str(row.get("Mã HH", row.get("Mã", ""))).strip()
                qty = float(row.get("Số Lượng", row.get("qty", 0)))
                
                # Ép về chữ thường và xóa khoảng trắng (Ví dụ: " Nhập " -> "nhập")
                t_type = str(row.get("Loại", "")).strip().lower()
                
                # Lọc chính xác: Nếu xóa phiếu Nhập -> trừ tồn. Xóa phiếu Xuất -> cộng tồn
                change = -qty if t_type == "nhập" else qty
                stock_changes[p_code] = stock_changes.get(p_code, 0) + change
            
            invalid_products = []
            for p_code, change in stock_changes.items():
                current_stock = stock_dict.get(p_code, 0)
                # Kiểm tra nghiêm ngặt
                if current_stock + change < 0:
                    invalid_products.append(f"• **{p_code}** (Tồn hiện tại: {current_stock:,.0f} ➔ Sau khi xóa sẽ thành: {current_stock + change:,.0f})")
            
            # ==========================================
            # BƯỚC 2: PHÂN NHÁNH XỬ LÝ
            # ==========================================
            if invalid_products:
                with msg_container:
                    st.error("🚫 **TỪ CHỐI XÓA: Giao dịch này sẽ làm tồn kho bị ÂM!**")
                    for msg in invalid_products:
                        st.write(msg)
            else:
                for index, row in selected_rows.iterrows():
                    p_code = str(row.get("Mã HH", row.get("Mã", ""))).strip()
                    qty = float(row.get("Số Lượng", row.get("qty", 0)))
                    # Giữ nguyên giá trị Loại gốc để truyền vào sheet
                    t_type_original = row.get("Loại", "")
                    
                    service.delete_transaction(index, p_code, qty, t_type_original)
                    
                st.cache_data.clear()
                st.success("🎉 Đã xóa và hoàn tác tồn kho thành công!")
                st.rerun()
                
    with col_no:
        if st.button("❌ Hủy bỏ"):
            st.rerun()

# --- CSS TỐI ƯU GIAO DIỆN KHÓA CỨNG TRÊN MOBILE ---
st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; padding-bottom: 11rem !important; }
    div.stButton > button[kind="primary"] { background-color: #28a745 !important; color: white !important; }
    h1 { padding-bottom: 0rem !important; margin-bottom: 0rem !important; }
    h3 { padding-top: 0rem !important; margin-top: 0rem !important; }
    div[data-testid="stSelectbox"] { margin-bottom: -1rem !important; }
            
    .stock-container {
        display: flex !important;
        justify-content: flex-end !important; /* Đẩy nội dung sang phải */
        align-items: center !important;
        margin-top: 0px !important;
        font-weight: bold !important;
        color: #28a745 !important;
        white-space: nowrap !important;
    }
    
    /* Ép chữ Loại và 2 nút Nhập/Xuất nằm ngang hàng tuyệt đối trên mọi màn hình (Cả PC lẫn Mobile) */
    div[data-testid="stRadio"] { 
        display: flex !important; 
        flex-direction: row !important; 
        align-items: center !important; 
        flex-wrap: nowrap !important;
        margin-bottom: -20px !important;
        margin-top: -20px !important; /* Cấm bẻ dòng */
    }
    div[data-testid="stRadio"] > label { 
        margin-bottom: 0px !important; 
        padding-bottom: 0px !important; 
        font-weight: bold !important; 
        font-size: 6px !important; 
        white-space: nowrap !important; /* Cấm chữ bị rớt xuống dưới */
        margin-right: 15px !important;
    }
    div[data-testid="stRadio"] > div { 
        display: flex !important; 
        flex-direction: row !important; 
        flex-wrap: nowrap !important; /* Cấm các nút Nhập/Xuất xếp chồng lên nhau */
    }
    /* Thêm đoạn này vào phần <style> ở đầu file */
    [data-testid="column"] {
        padding-left: 0px !important;
        padding-right: 0px !important;
    }

    /* Ép các cột sát nhau hơn bằng cách giảm margin mặc định */
    div[data-testid="stHorizontalBlock"] {
        gap: 10px !important;
    }        
    /* Bỏ khoảng trống thừa của AgGrid và các thành phần */
    .ag-header-cell-menu-button { display: none !important; }
    div[data-testid="stHorizontalBlock"] { gap: 5px !important; }
    div.stButton { margin-top: 0px !important; }
    
    /* Ép cột Tồn và Diễn giải nằm ngang trên điện thoại */
    .mobile-row { display: flex !important; flex-direction: row !important; align-items: center; gap: 10px; }
            
    /* Ẩn hoàn toàn menu AgGrid */
    .ag-header-cell-menu-button { display: none !important; }
    
    /* Ẩn thanh công cụ phía trên nếu có */
    .ag-menu { display: none !important; }
    
    /* Thu hẹp khoảng cách lưới lên trên */
    div[data-testid="stAgGrid"] { margin-top: -40px !important; }        
    </style>
""", unsafe_allow_html=True)

@st.cache_resource(show_spinner=False)
def get_data_service(): return DataService(mode="ONLINE")

service = get_data_service()
st.title("📦 Quản lý kho")

# --- QUẢN LÝ TRẠNG THÁI ĐĂNG NHẬP & MENU (CHỐNG MẤT KHI BẤM F5) ---
if "logged_in" not in st.session_state:
    if st.query_params.get("logged_in") == "true":
        st.session_state.logged_in = True
        st.session_state.user_name = st.query_params.get("user_name")
        st.session_state.user_role = st.query_params.get("user_role")
        st.session_state.current_menu = st.query_params.get("current_menu", "Danh mục hàng")
    else:
        st.session_state.logged_in = False
        st.session_state.user_name = None
        st.session_state.current_menu = "Danh mục hàng"

if not st.session_state.logged_in:
    with st.container(border=True):
        st.subheader("🔒 Đăng nhập hệ thống")
        user = st.text_input("Mã nhân viên (Username):")
        pwd = st.text_input("Mật khẩu:", type="password") 
        if st.button("Đăng nhập", type="primary", key="login_btn"):
            user_data = service.check_login(user, pwd)
            if user_data["status"]:
                st.session_state.logged_in = True
                st.session_state.user_name = user_data["name"]
                st.session_state.user_role = user_data["role"]
                
                st.query_params["logged_in"] = "true"
                st.query_params["user_name"] = user_data["name"]
                st.query_params["user_role"] = user_data["role"]
                st.query_params["current_menu"] = st.session_state.current_menu
                st.rerun() 
            else: 
                st.error("❌ Mã NV hoặc mật khẩu không đúng!")
    st.stop() 

# --- THANH SIDEBAR ẨN (CHỈ CHỨA THÔNG TIN USER & ĐĂNG XUẤT) ---
st.sidebar.write(f"👤 Người dùng: **{st.session_state.user_name}**")
st.sidebar.write(f"💼 Chức vụ: **{st.session_state.user_role}**") 
st.sidebar.markdown("---")

if st.sidebar.button("Đăng xuất", key="logout_btn"):
    st.session_state.logged_in = False
    st.query_params.clear() 
    st.rerun()

# --- ĐƯA MENU QUAY TRỞ LẠI MÀN HÌNH CHÍNH (ĐỂ KHÔNG BỊ MẤT) ---
menu_options = ["Danh mục hàng", "Nhập/Xuất Kho", "Báo cáo tồn kho", "Lịch sử giao dịch", "In phiếu xuất"]
if st.session_state.get("user_role") == "Quản lý":
    menu_options.append("Quản lý nhân viên")

if st.session_state.current_menu not in menu_options:
    st.session_state.current_menu = menu_options[0]

menu = st.selectbox(
    "Chức năng", 
    options=menu_options, 
    index=menu_options.index(st.session_state.current_menu),
    label_visibility="collapsed"
)

if menu != st.session_state.current_menu:
    st.session_state.current_menu = menu
    st.query_params["current_menu"] = menu 
    st.rerun()

# --- TAB 1: DANH MỤC HÀNG ---
if st.session_state.current_menu == "Danh mục hàng":
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
    
    c1, c2 = st.columns(2)
    with c1:
        with st.expander("➕ Thêm hàng hóa mới"):
            with st.form("add_form", clear_on_submit=True):
                code, name, unit = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
                if st.form_submit_button("Thêm hàng hóa"):
                    if not code or not name: st.warning("Nhập đủ Mã và Tên!")
                    elif service.check_product_exists(code.upper()): st.error("Mã đã tồn tại!")
                    else:
                        service.add_product(code.upper(), name, unit)
                        st.cache_data.clear(); st.success("Đã thêm thành công!"); st.rerun()
    with c2:
        with st.expander("🗑️ Xóa hàng hóa"):
            if products:
                # 1. Chuẩn bị dữ liệu
                df = pd.DataFrame(products, columns=["ID", "Mã", "Tên hàng hóa", "Đvt", "Tồn"])
                product_map = {f"{row['Mã']} - {row['Tên hàng hóa']}": row["Mã"] for _, row in df.iterrows()}
                
                # 2. Chọn sản phẩm
                selected_product = st.selectbox("Chọn hàng cần xóa", options=list(product_map.keys()), key="delete_product_select")
                del_code = product_map[selected_product]

                # 3. Nút xóa kích hoạt Dialog
                if st.button("🗑️ Xóa hàng này", use_container_width=True):
                    # Gọi dialog
                    confirm_delete(selected_product, del_code, service)

# --- TAB 2: NHẬP/XUẤT KHO ---
elif st.session_state.current_menu == "Nhập/Xuất Kho":
    st.subheader("🔄 Nhập/Xuất kho")
    
    trans_type = st.radio("Loại:", ["Nhập", "Xuất"], horizontal=True, key="trans_type")
    
    kho_nhap_list, kho_xuat_list = get_cached_config(service)
    products = get_cached_products(service)
    
    # Khởi tạo biến đếm key để reset ô số lượng
    if "qty_key" not in st.session_state:
        st.session_state.qty_key = 0
    
    if products:
        p_dict = {f"{p[1]} - {p[2]}": {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": p[4]} for p in products}
        selected = st.selectbox("Chọn hàng hóa", options=list(p_dict.keys()), index=None, key="product_select_field")
        
        # Chia 4 cột để gom nhóm
        c1, c2, c3, c4 = st.columns([0.8, 1, 1.5, 0.5])
        
        with c1: 
            # Sử dụng key động dựa trên st.session_state.qty_key
            qty = st.number_input("Số lượng", min_value=1.0, value=None, step=1.0, 
                                  key=f"qty_input_{st.session_state.qty_key}")
            
        with c2:
            if selected:
                current_stock = float(p_dict[selected]['Tồn'])
                unit = p_dict[selected]['Đvt']
                st.markdown(f"""
                    <div class='stock-container'>
                        Tồn: {current_stock:,.0f} {unit}
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.write("") 
                
        with c3: 
            note = st.selectbox("Diễn giải / Kho", options=(kho_nhap_list if trans_type == "Nhập" else kho_xuat_list), index=None, key="note_select_field")
            
        with c4:
            st.write("") # Căn chỉnh label
            st.write("") 
            
            if st.button("➕ Thêm hàng chờ", key="add_to_cart_btn"):
                if not selected or not qty or not note: 
                    st.warning("⚠️ Nhập đủ thông tin!")
                else:
                    if 'cart' not in st.session_state: st.session_state.cart = []
                    st.session_state.cart.append({
                        "Mã HH": p_dict[selected]["Mã"], 
                        "Tên HH": p_dict[selected]["Tên"], 
                        "Đvt": p_dict[selected]["Đvt"], 
                        "Số lượng": float(qty), 
                        "Ghi chú": note, 
                        "Loại": trans_type
                    })
                    
                    # TĂNG BIẾN ĐẾM ĐỂ RESET Ô SỐ LƯỢNG (Không gây lỗi)
                    st.session_state.qty_key += 1
                    st.rerun()

        # Phần hiển thị giỏ hàng và nút xác nhận
            if 'cart' not in st.session_state: st.session_state.cart = []
            if st.session_state.cart:
                # Lưới của bạn giữ nguyên
                edited_df_cart = st.data_editor(pd.DataFrame(st.session_state.cart), use_container_width=True, hide_index=True, key="cart_editor")
                
                # --- THÊM 2 NÚT BẤM CÙNG DÒNG Ở ĐÂY ---
                col_xac_nhan, col_huy = st.columns(2)
                
                with col_xac_nhan:
                    if st.button("✅ Xác nhận tất cả", type="primary", use_container_width=True, key="confirm_cart_btn"): 
                        for _, row in edited_df_cart.iterrows():
                            service.add_transaction(row["Mã HH"], row["Tên HH"], row["Số lượng"], row["Loại"], row["Ghi chú"], st.session_state.user_name)
                            service.update_stock(row["Mã HH"], row["Số lượng"], row["Loại"])
                        st.session_state.cart = []
                        st.cache_data.clear()
                        st.success(f"🎉 Giao dịch thành công!")
                        st.rerun()
                
                with col_huy:
                    if st.button("❌ Hủy hàng chờ", use_container_width=True, key="clear_cart_btn"):
                        st.session_state.cart = []
                        st.rerun()

# --- TAB 3: BÁO CÁO TỒN KHO ---
elif st.session_state.current_menu == "Báo cáo tồn kho":
    st.header("Báo cáo tồn kho")

# --- TAB 4: IN PHIẾU XUẤT ---
elif st.session_state.current_menu == "In phiếu xuất":
    show_print_export_view(service)

    show_report()

# --- TAB 5: LẠCH SỬ GIAO DỊCH ---
elif st.session_state.current_menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    
    # Chỉ cho phép Quản lý thực hiện xóa/sửa lịch sử để an toàn dữ liệu
    is_admin = st.session_state.get("user_role") == "Quản lý"
    
    history = get_cached_history(service)
    if history:
        # 1. Tạo DataFrame từ dữ liệu lịch sử
        # Đoạn này tự động khớp số lượng cột tùy theo cấu trúc dữ liệu trả về từ Sheet của bạn
        if len(history[0]) == 5:
            cols = ["Ngày", "Mã HH", "Loại", "Số Lượng", "Ghi Chú"]
        else:
            cols = ["Ngày", "Mã HH", "Tên hàng hóa", "Loại", "Số Lượng", "Ghi Chú", "Nhân viên"]
            
        df = pd.DataFrame(history, columns=cols)
        
        # 2. Thêm cột checkbox chọn dòng ở đầu bảng (chỉ hiện nếu là Quản lý)
        if is_admin:
            df.insert(0, "Chọn", False)
            
        # 3. Hiển thị bảng bằng st.data_editor (cho phép tích chọn)
        edited_df = st.data_editor(
            df, 
            key="history_editor", 
            use_container_width=True,
            hide_index=True,
            disabled=[col for col in df.columns if col != "Chọn"] # Khóa các cột khác, chỉ cho chọn checkbox
        )
        
        # 4. Xử lý nút bấm Xóa dành riêng cho Quản lý
        if is_admin:
            # Lọc ra các dòng được người dùng tích chọn ở cột "Chọn"
            selected_rows = edited_df[edited_df["Chọn"] == True]
            
            if not selected_rows.empty:
                # Nếu có chọn dòng, hiện nút xóa
                if st.button("🗑️ Xóa các giao dịch đã chọn", type="primary", use_container_width=True):
                    # Gọi hàm Dialog hỏi lại để chắc chắn
                    confirm_delete_history_dialog(selected_rows, service)
            else:
                # Nếu chưa chọn dòng nào, hiện nút mờ hướng dẫn
                st.button("💡 Hãy tích chọn dòng ở bảng trên để xóa", disabled=True, use_container_width=True)

# --- TAB 5: QUẢN LÝ NHÂN VIÊN ---
elif st.session_state.current_menu == "Quản lý nhân viên":
    if st.session_state.user_role != "Quản lý":
        st.error("🚫 Bạn không có quyền truy cập trang này!")
        st.stop()
    st.subheader("👥 Quản lý nhân viên")

    employees = get_cached_employees(service)
    if employees:
        df_emp = pd.DataFrame(employees, columns=["Mã NV", "Tên NV", "Số điện thoại", "Chức vụ", "Mật khẩu"])
        gb = GridOptionsBuilder.from_dataframe(df_emp)
        gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1)
        gb.configure_column("Mã NV", minWidth=80, editable=False, cellStyle={'textAlign': 'center'})
        gb.configure_column("Tên NV", minWidth=150, editable=True, cellStyle={'textAlign': 'left', 'backgroundColor': '#f0f8ff'}) 
        gb.configure_column("Số điện thoại", minWidth=120, editable=True, cellStyle={'textAlign': 'center', 'backgroundColor': '#f0f8ff'})
        gb.configure_column("Chức vụ", minWidth=120, editable=True, cellStyle={'textAlign': 'center', 'backgroundColor': '#f0f8ff'})
        gb.configure_column("Mật khẩu", minWidth=120, editable=True, cellStyle={'textAlign': 'center', 'backgroundColor': '#f0f8ff'})
        
        grid_response = AgGrid(df_emp, gridOptions=gb.build(), fit_columns_on_grid_load=True, theme='streamlit', update_mode=GridUpdateMode.MODEL_CHANGED, data_return_mode=DataReturnMode.FILTERED_AND_SORTED, height=400, key="employees_grid")
        edited_df_emp = pd.DataFrame(grid_response['data'])

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
            if st.button("💾 Lưu thay đổi", type="primary", key="save_emp_changes_btn"):
                for item in changes_to_save:
                    service.update_employee(item["Mã NV"], item["Tên NV"], item["Số điện thoại"], item["Chức vụ"])
                st.cache_data.clear()
                st.success("🎉 Đã cập nhật thông tin nhân viên thành công!")
                st.rerun()

    c1, c2 = st.columns(2)
    with c1:
        with st.expander("➕ Thêm nhân viên mới"):
            with st.form("add_emp_form", clear_on_submit=True):
                emp_code = st.text_input("Mã nhân viên (VD: NV01)")
                emp_name = st.text_input("Tên nhân viên")
                emp_phone = st.text_input("Số điện thoại")
                emp_role = st.selectbox("Chức vụ", ["Nhân viên kho", "Quản lý", "Kế toán", "Tài xế", "Khác"], key="add_emp_role_select")
                if st.form_submit_button("Thêm nhân viên"):
                    if not emp_code or not emp_name: st.warning("Vui lòng nhập đủ Mã và Tên nhân viên!")
                    elif service.check_employee_exists(emp_code.upper()): st.error("Mã nhân viên này đã tồn tại!")
                    else:
                        service.add_employee(emp_code.upper(), emp_name, emp_phone, emp_role)
                        st.cache_data.clear(); st.success("Đã thêm nhân viên thành công!"); st.rerun()
    with c2:
        with st.expander("🗑️ Xóa nhân viên"):
            if employees:
                del_emp_code = st.selectbox("Chọn mã NV cần xóa", options=df_emp["Mã NV"].tolist(), key="delete_emp_select")
                if st.button("Xác nhận xóa nhân viên", key="delete_emp_btn"):
                    service.delete_employee(del_emp_code)
                    st.cache_data.clear(); st.success(f"Đã xóa nhân viên {del_emp_code}!"); st.rerun()