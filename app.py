import streamlit as st
import pandas as pd
import io
import datetime
import streamlit.components.v1 as components  # Đã thêm dòng này để sửa lỗi sidebar mobile
from controllers.transaction_controller import TransactionController
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.print_export_view import show_print_export_view
from views.report_view_streamlit import show_report
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

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
                qty = float(row.get("Số lượng", row.get("Số Lượng", 0)))
                
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

# --- SCRIPT TỰ ĐỘNG ĐÓNG SIDEBAR TRÊN MOBILE ---
if st.session_state.get("force_close_sidebar", False):
    components.html(  # Đã sửa lỗi st.components.v1.html thành components.html
        """
        <script>
            // Tìm nút X (đóng sidebar) trên giao diện mobile của Streamlit và mô phỏng thao tác bấm
            var closeBtn = window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"]');
            if (closeBtn) {
                closeBtn.click();
            } else {
                // Dự phòng cho các phiên bản Streamlit cũ hơn
                var altBtn = window.parent.document.querySelector('button[aria-label="Close"]');
                if (altBtn) altBtn.click();
            }
        </script>
        """,
        height=0,
        width=0
    )
    # Tắt cờ đi để script không bị chạy lại vào các thao tác sau
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
                # 1. Kích hoạt trạng thái đăng nhập thành công
                st.session_state.logged_in = True
                st.session_state.user_name = user_data["name"]
                st.session_state.user_role = user_data["role"]
                
                # 2. ÉP BUỘC nhảy thẳng vào chức năng Danh mục hàng ngay lập tức
                st.session_state.current_menu = "Danh mục hàng"
                
                # 3. Lưu tham số vào URL trình duyệt để giữ trạng thái khi F5
                st.query_params["logged_in"] = "true"
                st.query_params["user_name"] = user_data["name"]
                st.query_params["user_role"] = user_data["role"]
                st.query_params["current_menu"] = "Danh mục hàng"
                st.session_state.force_close_sidebar = True
                # 4. Giải phóng giao diện, dẹp màn hình đăng nhập cũ ngay lập tức
                st.rerun() 
            else: 
                st.error("❌ Mã NV hoặc mật khẩu không đúng!")
                
    # Lệnh chặn này giữ người dùng ở lại form đăng nhập cho đến khi đăng nhập thành công
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
menu_options = ["Danh mục hàng", "Nhập/Xuất Kho", "Báo cáo tồn kho", "Lịch sử giao dịch", "In phiếu xuất"]
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
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên hàng hóa", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        gb = GridOptionsBuilder.from_dataframe(df[["Mã", "Tên hàng hóa", "Đvt", "Tồn"]])
        gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1)
        gb.configure_column("Mã", minWidth=50, editable=False)
        gb.configure_column("Tên hàng hóa", minWidth=150, editable=True)
        gb.configure_column("Đvt", minWidth=50, editable=True)
        gb.configure_column("Tồn", minWidth=60, editable=False, type=["numericColumn"], valueFormatter="Number(x).toLocaleString('en-US')")
        
        grid_response = AgGrid(
            df[["Mã", "Tên hàng hóa", "Đvt", "Tồn"]], 
            gridOptions=gb.build(), 
            fit_columns_on_grid_load=True, 
            theme='streamlit', 
            update_on=[{'event': 'cellValueChanged'}], 
            height=400
        )
    
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
                # 1. Chuẩn bị dữ liệu hiển thị: "Tên hàng hóa - (Tồn: X,XXX cái)"
                # Lưu ý: p[2] là Tên, p[3] là Đvt, p[4] là Tồn
                product_map = {
                    f"{row['Tên hàng hóa']} - (Tồn: {float(row['Tồn']):,.0f} {row['Đvt']})": row["Mã"] 
                    for _, row in df.iterrows()
                }
                
                # 2. Chọn sản phẩm từ map trên
                selected_product = st.selectbox(
                    "Chọn hàng cần xóa", 
                    options=list(product_map.keys()), 
                    key="delete_product_select"
                )
                del_code = product_map[selected_product]
                
                # 3. Lấy thông tin tồn kho hiện tại để kiểm tra ràng buộc
                current_stock = float(df[df["Mã"] == del_code]["Tồn"].iloc[0])

                # 4. Nút xóa kích hoạt Dialog (với ràng buộc đã sửa ở bước trước)
                if st.button("🗑️ Xóa hàng này", use_container_width=True):
                    # Kiểm tra lịch sử giao dịch
                    history = get_cached_history(service)
                    has_transaction = any(str(row[1]).strip() == str(del_code).strip() for row in history)
                    
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
    
    kho_nhap_list, kho_xuat_list = get_cached_config(service)
    products = get_cached_products(service)
    
    # Khởi tạo biến đếm key để reset ô số lượng
    if "qty_key" not in st.session_state:
        st.session_state.qty_key = 0
    
    if products:
        # --- ĐOẠN CODE ĐƯỢC TỐI ƯU HIỂN THỊ VÀ SẮP XẾP ---
        display_data = []
        p_dict = {}
        
        for p in products:
            # 1. Bắt lỗi an toàn cho dữ liệu Tồn kho (p[4])
            try:
                # Kiểm tra độ dài và loại bỏ khoảng trắng thừa, nếu lỗi thì gán bằng 0.0
                ton_kho = float(p[4]) if len(p) > 4 and str(p[4]).strip() != "" else 0.0
            except (ValueError, TypeError):
                ton_kho = 0.0
                
            # 2. Ráp chuỗi hiển thị
            ten_hien_thi = f"{p[2]} (Tồn: {ton_kho:,.0f} {p[3]})"
            display_data.append(ten_hien_thi)
            
            # 3. Lưu trữ toàn bộ thông tin gốc vào dict (Sử dụng luôn ton_kho đã làm sạch)
            p_dict[ten_hien_thi] = {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": ton_kho}
        
        # Sắp xếp danh sách hiển thị theo thứ tự Tên (A-Z)
        display_data.sort()
        # Hiển thị Selectbox với danh sách đã chuẩn hóa
        selected = st.selectbox("Chọn hàng hóa", options=display_data, index=None, key="product_select_field")
        # ---------------------------------------------------
        
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
                # Lưới hiển thị danh sách giỏ hàng
                edited_df_cart = st.data_editor(pd.DataFrame(st.session_state.cart), use_container_width=True, hide_index=True, key="cart_editor")
                
                # --- THÊM 2 NÚT BẤM CÙNG DÒNG Ở ĐÂY ---
                col_xac_nhan, col_huy = st.columns(2)
                
                with col_xac_nhan:
                    if st.button("✅ Xác nhận tất cả", type="primary", use_container_width=True, key="confirm_cart_btn"): 
    
                        # 1. Khởi tạo stock_dict TRƯỚC khi dùng
                        db_products = service.get_products()
                        stock_dict = {} # Khởi tạo mặc định
                        if db_products:
                            for p in db_products:
                                # p[1] là Mã (cột B), p[4] là Tồn (cột E)
                                p_code = str(p[1]).strip()
                                try:
                                    val = float(p[4]) if len(p) > 4 and p[4] else 0.0
                                except (ValueError, TypeError):
                                    val = 0.0
                                stock_dict[p_code] = val

                        # Bây giờ mới chạy vòng lặp kiểm tra xuất kho
                        error_msgs = []
                        for _, row in edited_df_cart.iterrows():
                            service.add_transaction(
                                row["Mã HH"], 
                                row["Tên HH"], 
                                row["Số lượng"], 
                                row["Loại"], 
                                row.get("Diễn giải", ""), 
                                st.session_state.user_name)
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

# --- TAB 5: LỊCH SỬ GIAO DỊCH (ĐÃ SỬA LỖI) ---
elif st.session_state.current_menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    
    # 1. Khởi tạo Controller an toàn
    if 't_controller' not in st.session_state:
        st.session_state.t_controller = TransactionController()
    t_controller = st.session_state.t_controller
    
    # 2. Lấy dữ liệu dạng DataFrame chuẩn
    history_df = t_controller.get_transaction_history()
    
    # 3. Kiểm tra DataFrame không rỗng
    if history_df is not None and not history_df.empty:
        # Nếu dữ liệu lấy về là list (cache cũ), chuyển sang DataFrame
        if not isinstance(history_df, pd.DataFrame):
            history_df = pd.DataFrame(history_df[1:], columns=history_df[0])
            
        # Thêm cột "Chọn" nếu là Quản lý
        is_admin = st.session_state.get("user_role") == "Quản lý"
        display_df = history_df.copy()
        if is_admin:
            display_df.insert(0, "Chọn", False)
            
        # Hiển thị bảng editor
        edited_df = st.data_editor(
            display_df, 
            use_container_width=True, 
            hide_index=True
        )
        
        # Xử lý xóa
        if is_admin:
            selected_rows = edited_df[edited_df["Chọn"] == True]
            if not selected_rows.empty:
                if st.button("🗑️ Xóa giao dịch đã chọn"):
                    confirm_delete_history_dialog(selected_rows, service)
    else:
        st.info("Chưa có dữ liệu lịch sử giao dịch.")

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
        
        # Đã cập nhật lại update_on thay cho update_mode
        grid_response = AgGrid(
            df_emp, 
            gridOptions=gb.build(), 
            fit_columns_on_grid_load=True, 
            theme='streamlit', 
            update_on=[{'event': 'cellValueChanged'}], 
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED, 
            height=400, 
            key="employees_grid"
        )
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
                    
# --- TAB 6: SAO LƯU DỮ LIỆU ---
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
                    # Đã thêm lệnh cắt 5 cột đầu tiên để đảm bảo lúc xuất Excel không bị lỗi
                    products_data = [row[:5] for row in products]
                    pd.DataFrame(products_data, columns=["ID", "Mã", "Tên hàng hóa", "Đvt", "Tồn"]).to_excel(writer, index=False, sheet_name="Danh_Muc_Ton")
                
                if history:
                    # Lấy số cột thực tế và gán tiêu đề tương ứng để xuất Excel không bị lỗi
                    num_cols_hist = len(history[0])
                    
                    if num_cols_hist == 5:
                        cols_hist = ["Ngày tháng", "Mã HH", "Loại", "Số lượng", "Diễn giải"]
                    elif num_cols_hist == 6:
                        cols_hist = ["Ngày tháng", "Mã HH", "Tên hàng hóa", "Loại", "Số lượng", "Diễn giải"]
                    elif num_cols_hist == 7:
                        cols_hist = ["Ngày tháng", "Mã HH", "Tên hàng hóa", "Loại", "Số lượng", "Diễn giải", "Nhân Viên"]
                    elif num_cols_hist == 8:
                        cols_hist = ["Ngày tháng", "Mã HH", "Tên hàng hóa", "Đvt", "Loại", "Số lượng", "Diễn giải", "Nhân Viên"]
                    else:
                        cols_hist = [f"Cột {i+1}" for i in range(num_cols_hist)]
                        
                    pd.DataFrame(history, columns=cols_hist).to_excel(writer, index=False, sheet_name="Lich_Su_Giao_Dich")
                
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

