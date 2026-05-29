import streamlit as st
import pandas as pd
import io
import datetime
import streamlit.components.v1 as components  
from controllers.transaction_controller import TransactionController
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.print_export_view import show_print_export_view
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from views.report_view_streamlit import show_report, export_history_to_excel, get_count_import_by_department
from datetime import date

if "user_role" not in st.session_state:
    st.session_state.user_role = None

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
        if st.button("✅ Xác nhận xóa XXX"):
            
            # BƯỚC 1: KIỂM TRA RÀNG BUỘC (SIÊU CHUẨN XÁC)
            products = service.get_products()
            stock_dict = {str(p[1]).strip(): float(p[4]) for p in products} if products else {}
            
            stock_changes = {}
            for index, row in selected_rows.iterrows():
                p_code = str(row.get("Mã HH", row.get("Mã", ""))).strip()
                qty = float(row.get("Số lượng", row.get("Số Lượng", 0)))
                
                t_type = str(row.get("Loại", "")).strip().lower()
                change = -qty if t_type == "nhập" else qty
                stock_changes[p_code] = stock_changes.get(p_code, 0) + change
            
            invalid_products = []
            for p_code, change in stock_changes.items():
                current_stock = stock_dict.get(p_code, 0)
                if current_stock + change < 0:
                    invalid_products.append(f"• **{p_code}** (Tồn hiện tại: {current_stock:,.0f} ➔ Sau khi xóa sẽ thành: {current_stock + change:,.0f})")
            
            # BƯỚC 2: PHÂN NHÁNH XỬ LÝ
            if invalid_products:
                with msg_container:
                    st.error("🚫 **TỪ CHỐI XÓA: Giao dịch này sẽ làm tồn kho bị ÂM!**")
                    for msg in invalid_products:
                        st.write(msg)
            else:
                for index, row in selected_rows.iterrows():
                    p_code = str(row.get("Mã HH", row.get("Mã", ""))).strip()
                    qty = float(row.get("Số Lượng", row.get("qty", 0)))
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

# --- SCRIPT TỰ ĐỘNG ĐÓNG SIDEBAR TRÊN MOBILE ---
if st.session_state.get("force_close_sidebar", False):
    components.html(
        """
        <script>
            var closeBtn = window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"]');
            if (closeBtn) {
                closeBtn.click();
            } else {
                var altBtn = window.parent.document.querySelector('button[aria-label="Close"]');
                if (altBtn) altBtn.click();
            }
        </script>
        """,
        height=0,
        width=0
    )
    st.session_state.force_close_sidebar = False

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

# --- KHÓA MÀN HÌNH ĐĂNG NHẬP NẾU CHƯA LOGGED_IN ---
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
                
                st.session_state.current_menu = "Danh mục hàng"
                
                st.query_params["logged_in"] = "true"
                st.query_params["user_name"] = user_data["name"]
                st.query_params["user_role"] = user_data["role"]
                st.query_params["current_menu"] = "Danh mục hàng"
                st.session_state.force_close_sidebar = True
                st.rerun() 
            else: 
                st.error("❌ Mã NV hoặc mật khẩu không đúng!")
                
    st.stop() 

# --- THANH SIDEBAR ẨN (CHỈ CHỨA THÔNG TIN USER & ĐĂNG XUẤT) ---
user_name = st.session_state.get("user_name", "Khách")
user_role = st.session_state.get("user_role", "Chưa đăng nhập")
st.sidebar.write(f"👤 Người dùng: **{st.session_state.user_name}**")
st.sidebar.write(f"💼 Chức vụ: **{st.session_state.user_role}**") 
st.sidebar.markdown("---")

if st.sidebar.button("Đăng xuất", key="logout_btn"):
    st.session_state.logged_in = False
    st.query_params.clear() 
    st.rerun()

# --- ĐƯA MENU QUAY TRỞ LẠI MÀN HÌNH CHÍNH (ĐỂ KHÔNG BỊ MẤT) ---
menu_options = ["Danh mục hàng", "Nhập/Xuất Kho", "Báo cáo tồn kho", "Lịch sử giao dịch", "In phiếu xuất", "Thống kê nhập kho"]
if st.session_state.get("user_role") == "Quản lý":
    menu_options.append("Quản lý nhân viên")
    menu_options.append("Sao lưu dữ liệu")

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
        # 1. THÊM CỘT "Mức tối thiểu" VÀ "Ghi chú" VÀO DATAFRAME
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên hàng hóa", "Đvt", "Tồn", "Nhóm", "Mức tối thiểu", "Ghi chú"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        df["Mức tối thiểu"] = pd.to_numeric(df["Mức tối thiểu"], errors="coerce").fillna(0)

       # ==============================================================
        # CẤU HÌNH ĐỘ RỘNG CỘT VÀ CHỨC NĂNG SỬA CHO BẢNG DANH MỤC HÀNG
        # ==============================================================
        gb = GridOptionsBuilder.from_dataframe(df)
        
        gb.configure_default_column(sortable=True, filter=True, resizable=True, editable=True)
        gb.configure_column("ID", hide=True)
        gb.configure_column("Mã", minWidth=80, maxWidth=100, editable=False, cellStyle={'textAlign': 'center'})
        gb.configure_column("Tên hàng hóa", minWidth=200, flex=1, cellStyle={'textAlign': 'left'}) 
        gb.configure_column("Đvt", minWidth=60, maxWidth=90, cellStyle={'textAlign': 'center'})
        gb.configure_column("Tồn", minWidth=80, maxWidth=120, editable=False, cellStyle={'textAlign': 'right', 'fontWeight': 'bold', 'color': '#28a745'})
        gb.configure_column("Nhóm", minWidth=120, maxWidth=150)
        gb.configure_column("Mức tối thiểu", minWidth=120, maxWidth=150, type=["numericColumn"], valueFormatter="data['Mức tối thiểu'].toFixed(0)", cellStyle={'textAlign': 'right'})
        gb.configure_column("Ghi chú", minWidth=150, editable=True, cellStyle={'textAlign': 'left'})

        grid_response = AgGrid(
            df, 
            gridOptions=gb.build(), 
            fit_columns_on_grid_load=True, 
            theme='streamlit', 
            update_mode='MODEL_CHANGED', 
            data_return_mode='AS_INPUT', 
            height=400,
            key="products_grid") 
        # ==============================================================
        # 2. KHỐI LOGIC QUÉT THAY ĐỔI (Đã Tối Ưu Tốc Độ Siêu Nhanh)
        # ==============================================================
        changes_to_save = [] # BẮT BUỘC KHỞI TẠO Ở NGOÀI CÙNG
        
        if grid_response['data'] is not None:
            edited_df = pd.DataFrame(grid_response['data'])
            
            def to_float(val):
                try: return float(val)
                except: return 0.0

            if not edited_df.empty and not df.empty:
                # [BÍ QUYẾT TĂNG TỐC TẠI ĐÂY] 
                # Chuyển DataFrame gốc thành Dictionary để tra cứu tức thì
                df_clean = df.copy()
                df_clean['Mã'] = df_clean['Mã'].astype(str).str.strip()
                
                # --- MÀNG LỌC CHỐNG LỖI ---
                # 1. Loại bỏ các dòng rỗng (Mã trống do Google Sheets dư dòng)
                df_clean = df_clean[df_clean['Mã'] != ""]
                # 2. Xử lý trùng lặp: Nếu có 2 mã giống hệt nhau do nhập nhầm, giữ lại dòng đầu tiên
                df_clean = df_clean.drop_duplicates(subset=['Mã'], keep='first')
                # --------------------------

                orig_dict = df_clean.set_index('Mã').to_dict('index')

                for i in range(len(edited_df)):
                    ma = str(edited_df.iloc[i]["Mã"]).strip()
                    
                    # Tra cứu trực tiếp từ Dictionary (O(1))
                    if ma not in orig_dict: 
                        continue
                    orig_row = orig_dict[ma]
                    
                    ten_moi = str(edited_df.iloc[i].get("Tên hàng hóa", "")).strip()
                    dvt_moi = str(edited_df.iloc[i].get("Đvt", "")).strip()
                    nhom_moi = str(edited_df.iloc[i].get("Nhóm", "")).strip()
                    muc_moi = to_float(edited_df.iloc[i].get("Mức tối thiểu", 0))
                    ghi_chu_moi = str(edited_df.iloc[i].get("Ghi chú", "")).strip()
                    
                    ten_cu = str(orig_row.get("Tên hàng hóa", "")).strip()
                    dvt_cu = str(orig_row.get("Đvt", "")).strip()
                    nhom_cu = str(orig_row.get("Nhóm", "")).strip()
                    muc_cu = to_float(orig_row.get("Mức tối thiểu", 0))
                    ghi_chu_cu = str(orig_row.get("Ghi chú", "")).strip()
                    
                    # LOGIC SO SÁNH
                    is_changed = (
                        ten_moi != ten_cu or 
                        dvt_moi != dvt_cu or 
                        nhom_moi != nhom_cu or 
                        abs(muc_moi - muc_cu) > 0.001 or
                        ghi_chu_moi != ghi_chu_cu
                    )
                    
                    if is_changed:
                        changes_to_save.append({
                            "Mã": ma, "Tên": ten_moi, "Đvt": dvt_moi, 
                            "Nhóm": nhom_moi, "Mức": muc_moi, "Ghi chú": ghi_chu_moi
                        })

        # ==============================================================
        # HIỆN THÔNG BÁO VÀ NÚT BẤM (CHỈ KHI CÓ THAY ĐỔI > 0)
        # ==============================================================
        if len(changes_to_save) > 0:
            st.markdown("### Có thay đổi cần lưu")
            st.warning(f"⚠️ Có {len(changes_to_save)} hàng hóa đã thay đổi thông tin!")
            
            if st.button("💾 Lưu thay đổi vào Google Sheets", type="primary"):
                with st.spinner("Đang cập nhật..."):
                    for item in changes_to_save:
                        service.update_product(item["Mã"], item["Tên"], item["Đvt"], item["Nhóm"], item["Mức"], item["Ghi chú"])
                
                st.cache_data.clear()
                st.success("🎉 Cập nhật thành công!")
                st.rerun()
    
    c1, c2 = st.columns(2)
    with c1:
        with st.expander("➕ Thêm hàng hóa mới"):
            with st.form("add_form", clear_on_submit=True):
                code = st.text_input("Mã hàng")
                name = st.text_input("Tên hàng")
                unit = st.text_input("Đơn vị tính")
                danh_sach_nhom = ["Vật tư", "Dụng cụ sản xuất", "Phụ gia", "Bao bì", "PE", "Khác"]
                group = st.selectbox("Chọn nhóm hàng", danh_sach_nhom)
                min_stock = st.number_input("Mức tồn tối thiểu để cảnh báo", min_value=0, value=10, step=1)
                note = st.text_input("Ghi chú")
                
                if st.form_submit_button("Thêm hàng hóa"):
                    if not code or not name: 
                        st.warning("Nhập đủ Mã và Tên!")
                    elif service.check_product_exists(code.upper()): 
                        st.error("Mã đã tồn tại!")
                    else:
                        service.add_product(code.upper(), name, unit, group, min_stock, note)
                        st.cache_data.clear()
                        st.success("Đã thêm thành công!")
                        st.rerun()
    with c2:
        with st.expander("🗑️ Xóa hàng hóa"):
            if products:
                # [TỐI ƯU TỐC ĐỘ 2]: Dùng hàm Zip thay vì iterrows() để nạp danh sách vào Selectbox cực nhanh
                product_map = {
                    f"{n} - (Tồn: {float(t):,.0f} {d})": m
                    for m, n, t, d in zip(df["Mã"], df["Tên hàng hóa"], df["Tồn"], df["Đvt"])
                }
                
                selected_product = st.selectbox(
                    "Chọn hàng cần xóa", 
                    options=list(product_map.keys()), 
                    key="delete_product_select"
                )
                del_code = product_map[selected_product]
                current_stock = float(df[df["Mã"] == del_code]["Tồn"].iloc[0])

                if st.button("🗑️ Xóa hàng này", use_container_width=True):
                    history = get_cached_history(service)
                    
                    has_transaction = False
                    if isinstance(history, pd.DataFrame) and not history.empty:
                        col_ma = "Mã HH" if "Mã HH" in history.columns else "Mã"
                        if col_ma in history.columns:
                            has_transaction = str(del_code).strip() in history[col_ma].astype(str).str.strip().values
                    
                    if current_stock != 0:
                        st.error(f"🚫 Không thể xóa: Hàng hóa này vẫn còn tồn kho ({current_stock:,.0f})!")
                    elif has_transaction:
                        st.error("🚫 Không thể xóa: Hàng hóa này đã có lịch sử giao dịch!")
                    else:
                        confirm_delete(selected_product, del_code, service)

# --- TAB 2: NHẬP/XUẤT KHO ---
elif st.session_state.current_menu == "Nhập/Xuất Kho":
    st.subheader("🔄 Nhập/Xuất kho")
    
    trans_type = st.radio("Loại:", ["Nhập", "Xuất"], horizontal=True, key="trans_type")
    
    # 1. Bổ sung nhận 3 danh sách từ hàm get_cached_config
    kho_nhap_list, kho_xuat_list, bo_phan_list = get_cached_config(service)
    products = get_cached_products(service)
    
    if "qty_key" not in st.session_state:
        st.session_state.qty_key = 0
    
    if products:
        display_data = []
        p_dict = {}
        
        for p in products:
            try:
                ton_kho = float(p[4]) if len(p) > 4 and str(p[4]).strip() != "" else 0.0
            except (ValueError, TypeError):
                ton_kho = 0.0
                
            ten_hien_thi = f"{p[2]} (Tồn: {ton_kho:,.0f} {p[3]})"
            display_data.append(ten_hien_thi)
            
            p_dict[ten_hien_thi] = {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": ton_kho}
        
        display_data.sort()
        selected = st.selectbox("Chọn hàng hóa", options=display_data, index=None, key="product_select_field")
        
        # 2. Chia lại thành 5 cột để có chỗ hiển thị chọn "Bộ phận"
        c1, c2, c3, c4, c5 = st.columns([0.8, 1, 1.2, 1.2, 0.8])
        
        with c1: 
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
            # 3. Chỉ hiển thị chọn Bộ phận nếu là loại "Nhập"
            if trans_type == "Nhập":
                bo_phan = st.selectbox("Bộ phận", options=bo_phan_list, index=None, key="bophan_select_field")
            else:
                bo_phan = ""
                st.write("") 
                
        with c5:
            st.write("") 
            st.write("") 
            
            if st.button("➕ Thêm", key="add_to_cart_btn"):
                # Ràng buộc: Nhập kho thì bắt buộc phải có Bộ phận
                if not selected or not qty or not note or (trans_type == "Nhập" and not bo_phan): 
                    st.warning("⚠️ Nhập đủ thông tin!")
                else:
                    if 'cart' not in st.session_state: st.session_state.cart = []
                    st.session_state.cart.append({
                        "Mã HH": p_dict[selected]["Mã"], 
                        "Tên HH": p_dict[selected]["Tên"], 
                        "Đvt": p_dict[selected]["Đvt"], 
                        "Số lượng": float(qty), 
                        "Diễn Giải": note, 
                        "Bộ phận": bo_phan, # Lưu giá trị bộ phận vào giỏ hàng
                        "Loại": trans_type
                    })
                    
                    st.session_state.qty_key += 1
                    st.rerun()

        if 'cart' not in st.session_state: st.session_state.cart = []
        if st.session_state.cart:
            edited_df_cart = st.data_editor(pd.DataFrame(st.session_state.cart), use_container_width=True, hide_index=True, key="cart_editor")
            
            col_xac_nhan, col_huy = st.columns(2)
            
            with col_xac_nhan:
                if st.button("✅ Xác nhận tất cả", type="primary", use_container_width=True, key="confirm_cart_btn"): 
    
                    db_products = service.get_products()
                    stock_dict = {} 
                    if db_products:
                        for p in db_products:
                            p_code = str(p[1]).strip()
                            try:
                                val = float(p[4]) if len(p) > 4 and p[4] else 0.0
                            except (ValueError, TypeError):
                                val = 0.0
                            stock_dict[p_code] = val

                    error_msgs = []
                    for _, row in edited_df_cart.iterrows():
                        # 4. Gửi giá trị Bộ phận xuống hàm add_transaction
                        service.add_transaction(
                            row["Mã HH"], 
                            row["Tên HH"], 
                            row["Số lượng"], 
                            row["Loại"], 
                            row.get("Diễn Giải", ""), 
                            st.session_state.user_name,
                            row.get("Bộ phận", "")
                        )
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
    show_report()

# --- TAB 4: IN PHIẾU XUẤT ---
elif st.session_state.current_menu == "In phiếu xuất":
    show_print_export_view(service)

# --- TAB MỚI: THỐNG KÊ NHẬP KHO ---
elif st.session_state.current_menu == "Thống kê nhập kho":
    st.header("📊 Thống kê tần suất nhập kho")
    
    col1, col2 = st.columns(2)
    with col1:
        s_date = st.date_input("Từ ngày", key="rep_s")
    with col2:
        e_date = st.date_input("Đến ngày", key="rep_e")
        
    if st.button("Xem thống kê", type="primary"):
        # Gọi hàm tính toán
        df_freq = get_count_import_by_department(service, s_date, e_date)
        
        if not df_freq.empty:
            st.table(df_freq)
            
            # Xuất Excel cho bảng thống kê này
            st.download_button(
                label="📥 Xuất thống kê ra Excel",
                data=export_history_to_excel(df_freq), # Tái sử dụng hàm xuất đã có
                file_name=f"ThongKeNhapKho_{s_date.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Không có dữ liệu nhập kho trong khoảng thời gian này.")

# --- TAB 5: LỊCH SỬ GIAO DỊCH ---
elif st.session_state.current_menu == "Lịch sử giao dịch":
    st.subheader("📜 Lịch sử giao dịch")
    
    if 't_controller' not in st.session_state:
        st.session_state.t_controller = TransactionController()
    t_controller = st.session_state.t_controller
    
    history_df = t_controller.get_transaction_history()
    
    if history_df is not None and not history_df.empty:
        if not isinstance(history_df, pd.DataFrame):
            history_df = pd.DataFrame(history_df[1:], columns=history_df[0])
            
        # --- BỔ SUNG: NÚT XUẤT EXCEL Ở ĐÂY ---
        st.download_button(
            label="📥 Xuất lịch sử giao dịch ra Excel",
            data=export_history_to_excel(history_df),
            file_name=f"LichSuGiaoDich_{date.today().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
            
        is_admin = st.session_state.get("user_role") == "Quản lý"
        display_df = history_df.copy()
        if is_admin:
            display_df.insert(0, "Chọn", False)
            
        edited_df = st.data_editor(
            display_df, 
            use_container_width=True, 
            hide_index=True
        )
        
        if is_admin:
            selected_rows = edited_df[edited_df["Chọn"] == True]
            if not selected_rows.empty:
                if st.button("🗑️ Xóa giao dịch đã chọn"):
                    confirm_delete_history_dialog(selected_rows, service)
    else:
        st.info("Chưa có dữ liệu lịch sử giao dịch.")

# --- TAB 6: QUẢN LÝ NHÂN VIÊN ---
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
        
        grid_response = AgGrid(
            df_emp, 
            gridOptions=gb.build(), 
            fit_columns_on_grid_load=True, 
            theme='streamlit', 
            update_on=[{'event': 'cellValueChanged'}], 
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED, 
            height=400, 
            key="employees_grid")
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
                    
# --- TAB 7: SAO LƯU DỮ LIỆU ---
elif st.session_state.current_menu == "Sao lưu dữ liệu":
    if st.session_state.user_role != "Quản lý":
        st.error("🚫 Bạn không có quyền truy cập trang này!")
        st.stop()
        
    st.subheader("💾 Sao lưu an toàn dữ liệu")
    st.info("Dữ liệu gốc đang được bảo vệ an toàn trên Google. Tuy nhiên, bạn có thể tải thêm một bản sao lưu ngoại tuyến (Offline) về máy tính để phòng hờ.")
    
    if st.button("📦 Tạo bản sao lưu ngay", type="primary"):
        with st.spinner("Đang tổng hợp dữ liệu..."):
            products = get_cached_products(service)
            history = get_cached_history(service)
            employees = get_cached_employees(service)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if products:
                    products_data = [row[:5] for row in products]
                    pd.DataFrame(products_data, columns=["ID", "Mã", "Tên hàng hóa", "Đvt", "Tồn"]).to_excel(writer, index=False, sheet_name="Danh_Muc_Ton")
                
                if history is not None and not history.empty: # history là Dataframe
                    history.to_excel(writer, index=False, sheet_name="Lich_Su_Giao_Dich")
                
                if employees:
                    pd.DataFrame(employees, columns=["Mã NV", "Tên NV", "SĐT", "Chức vụ", "Mật khẩu"]).to_excel(writer, index=False, sheet_name="Nhan_Vien")
            
            today_str = datetime.datetime.now().strftime("%d_%m_%Y")
            st.success("🎉 Tạo file thành công! Hãy bấm nút bên dưới để tải về máy.")
            st.download_button(
                label="⬇️ Tải file Excel sao lưu",
                data=output.getvalue(),
                file_name=f"Backup_Kho_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )