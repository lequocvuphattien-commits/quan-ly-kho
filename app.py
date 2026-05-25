import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report
# --- [QUAN TRỌNG]: Đã thêm GridUpdateMode vào thư viện import ---
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode

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

# Cấu hình trang (Luôn để đầu tiên)
st.set_page_config(page_title="Quản Lý Kho", layout="wide", initial_sidebar_state="collapsed")

# CSS tinh chỉnh màu sắc nút bấm và giao diện
st.markdown("""
    <style>
    /* 1. KÉO NỘI DUNG LÊN SÁT CẠNH TRÊN MÀN HÌNH */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
    /* Làm nổi bật nút Xác nhận tất cả (Xanh lá) */
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
    }    
    
    /* 2. KÉO GẦN TIÊU ĐỀ CHÍNH VÀ MENU LẠI VỚI NHAU */
    h1 {
        padding-bottom: 0rem !important;
        margin-bottom: 0rem !important; 
    }   
    
    /* 3. KÉO NỘI DUNG (SUBHEADER) LÊN SÁT MENU */
    h3 {
        padding-top: 0rem !important;
        margin-top: 0rem !important; /* Lực hút kéo phần nội dung bên dưới trồi lên */
    }
    
    /* Ép khoảng trống dưới menu nhỏ lại */
    div[data-testid="stSelectbox"] {
        margin-bottom: -1rem !important; 
    }  
    </style>
""", unsafe_allow_html=True)

# --- [TỐI ƯU]: KHỞI TẠO KẾT NỐI API 1 LẦN DUY NHẤT ---
@st.cache_resource(show_spinner=False)
def get_data_service():
    # Ứng dụng chỉ gọi lên Google 1 lần lúc mới mở app để lấy token kết nối
    return DataService(mode="ONLINE")

service = get_data_service()

st.title("📦 Quản lý kho")

# --- BẢO MẬT: MÀN HÌNH ĐĂNG NHẬP ---
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    # Vẽ một khung viền cho form đăng nhập nhìn chuyên nghiệp hơn
    with st.container(border=True):
        st.subheader("🔒 Đăng nhập hệ thống")
        st.info("Vui lòng nhập mật khẩu để truy cập phần mềm quản lý.")
        
        pwd = st.text_input("Mật khẩu:", type="password") 
        
        if st.button("Đăng nhập", type="primary"):
            if pwd == "123":  
                st.session_state.logged_in = True
                st.rerun() 
            else:
                st.error("❌ Mật khẩu không đúng!")
                
    st.stop() 
# ==========================================

# Đưa menu ra màn hình chính, bỏ chữ "sidebar." đi
menu = st.selectbox(
    "Chức năng", 
    ["Danh mục hàng", "Nhập/Xuất Kho", "Báo cáo tồn kho", "Lịch sử giao dịch"],
    label_visibility="collapsed" 
)

# --- TAB 1: DANH MỤC HÀNG ---
if menu == "Danh mục hàng":
    st.subheader("📋 Danh mục hàng")
    products = get_cached_products(service)
    
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)

        # --- [THAY ĐỔI MỚI]: CHUYỂN SANG DÙNG AGGRID ĐỂ CÓ BỘ LỌC CHUYÊN NGHIỆP ---
        #st.markdown("💡 *Mẹo: Bấm vào biểu tượng 3 gạch trên tiêu đề cột để **Lọc**. Click đúp vào ô có nền xanh (Tên, Đvt) để **Sửa**.*")
        
        gb = GridOptionsBuilder.from_dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]])
        gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1) # Bật bộ lọc cho toàn bộ các cột
        
        # Cấu hình chi tiết từng cột (Chỉ cho phép sửa Tên và Đvt, tô màu nền xanh nhạt để phân biệt)
        gb.configure_column("Mã", minWidth=100, editable=False, cellStyle={'textAlign': 'center'})
        gb.configure_column("Tên", minWidth=200, editable=True, cellStyle={'backgroundColor': '#f0f8ff'}) 
        gb.configure_column("Đvt", minWidth=100, editable=True, cellStyle={'backgroundColor': '#f0f8ff'})
        gb.configure_column("Tồn", minWidth=100, editable=False, type=["numericColumn"], valueFormatter="Number(x).toLocaleString('en-US')", cellStyle={'textAlign': 'right'})
        
        go = gb.build()
        
        grid_response = AgGrid(
            df[["Mã", "Tên", "Đvt", "Tồn"]],
            gridOptions=go,
            fit_columns_on_grid_load=True,
            theme='streamlit',
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED, # Khi xuất excel sẽ chỉ xuất những dòng đã lọc
            height=400
        )

        edited_df = pd.DataFrame(grid_response['data'])

        # --- XỬ LÝ LƯU THAY ĐỔI TỪ AGGRID ---
        has_changes = False
        changes_to_save = [] 
        
        if not edited_df.empty:
            for i in range(len(edited_df)):
                ma = edited_df.iloc[i]["Mã"]
                ten_moi = edited_df.iloc[i]["Tên"]
                dvt_moi = edited_df.iloc[i]["Đvt"]
                
                # Tìm dòng dữ liệu gốc bằng Mã để so sánh độ chênh lệch
                orig_row = df[df["Mã"] == ma].iloc[0]
                
                if ten_moi != orig_row["Tên"] or dvt_moi != orig_row["Đvt"]:
                    has_changes = True
                    changes_to_save.append({"Mã": ma, "Tên": ten_moi, "Đvt": dvt_moi})

        if has_changes:
            st.info("⚠️ Có thay đổi chưa được lưu!")
            if st.button("💾 Lưu thay đổi", type="primary"):
                for item in changes_to_save:
                    service.update_product(item["Mã"], item["Tên"], item["Đvt"])
                
                st.cache_data.clear()
                st.success("🎉 Đã cập nhật thông tin thành công!")
                st.rerun()

        # Xuất file Excel (Dựa trên dữ liệu đã lọc trên lưới AgGrid)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if not edited_df.empty:
                edited_df.to_excel(writer, index=False, sheet_name='DanhMuc')
                worksheet = writer.sheets['DanhMuc']
                worksheet.freeze_panes = 'A2'
                max_row = worksheet.max_row
                max_col = worksheet.max_column
                if max_row > 1:
                    worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
                for col in range(1, max_col + 1):
                    worksheet.column_dimensions[get_column_letter(col)].width = 15
        
        st.download_button(
            label="📥 Xuất danh mục ra Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="DanhMucHangHoa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # --- GIAO DIỆN THÊM VÀ XÓA HÀNG HÓA NẰM CẠNH NHAU ---
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
                        st.cache_data.clear()
                        st.success("Đã thêm thành công!"); st.rerun()
    with c2:
        with st.expander("🗑️ Xóa hàng hóa"):
            if products:
                del_code = st.selectbox("Chọn mã hàng cần xóa", options=df["Mã"].tolist())
                if st.button("Xác nhận xóa"):
                    service.delete_product(del_code)
                    st.cache_data.clear()
                    st.success(f"Đã xóa {del_code}!")
                    st.rerun()

# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất Kho":
    st.subheader("🔄 Nhập/Xuất kho")
    
    trans_type = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
    kho_nhap_list, kho_xuat_list = service.get_config_options()
    products = get_cached_products(service)
    
    if products:
        p_dict = {
            f"{p[1]} - {p[2]} (Tồn: {float(p[4]):,.0f} {p[3]})": {
                "Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": p[4]
            } 
            for p in products
        }
        
        with st.container(border=True):
            selected = st.selectbox(
                "Chọn hàng hóa", 
                options=list(p_dict.keys()), 
                index=None, 
                placeholder="🔍 Gõ tìm kiếm mã hoặc tên hàng..."
            )
            
            c1, c2, c3 = st.columns([1.5, 2, 1])
            with c1:
                qty = st.number_input("Số lượng", min_value=1.0, value=None, step=1.0, format="%.0f", placeholder="Nhập số...")
            with c2:
                current_options = kho_nhap_list if trans_type == "Nhập" else kho_xuat_list
                note = st.selectbox("Diễn giải / Kho", options=current_options, index=None, placeholder="Chọn địa điểm...")
            with c3:
                st.write("") 
                if st.button("➕ Thêm vào lưới", key="add_to_cart"):
                    if not selected or not qty or not note:
                        st.warning("⚠️ Vui lòng chọn hàng hóa, số lượng và địa điểm!")
                    else:
                        prod_data = p_dict[selected]
                        if trans_type == "Xuất" and qty > float(prod_data["Tồn"]):
                            st.error("❌ Không đủ tồn kho!")
                            st.stop()
                        st.session_state.cart.append({
                            "Mã HH": prod_data["Mã"], "Tên HH": prod_data["Tên"], "Đvt": prod_data["Đvt"],
                            "Số lượng": float(qty), "Ghi chú": note, "Loại": trans_type
                        })
                        st.rerun()
                    
        with st.expander(f"➕ Thêm địa điểm mới: {trans_type}"):
            new_kho = st.text_input("Tên địa điểm mới", placeholder="Ví dụ: Kho hoặc địa điểm...")
            if st.button("Lưu địa điểm mới"):
                if new_kho:
                    service.add_config_option(trans_type, new_kho)
                    st.success(f"Đã thêm {new_kho} vào danh mục {trans_type}!")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.warning("Vui lòng nhập tên kho!")

        if 'cart' not in st.session_state: st.session_state.cart = []

        if st.session_state.cart:
            st.markdown("### 📋 Lưới chờ xử lý")
            edited_df_cart = st.data_editor(
                pd.DataFrame(st.session_state.cart),
                column_config={
                    "Mã HH": st.column_config.TextColumn("Mã HH", disabled=True),
                    "Tên HH": st.column_config.TextColumn("Tên HH", disabled=True),
                    "Đvt": st.column_config.TextColumn("Đvt", disabled=True),
                    "Loại": st.column_config.TextColumn("Loại", disabled=True),
                    "Số lượng": st.column_config.NumberColumn("Số lượng", required=True, min_value=1.0, format="%.0f"),
                    "Ghi chú": st.column_config.TextColumn("Ghi chú")
                },
                hide_index=True, use_container_width=True
            )
            
            if st.button("✅ Xác nhận tất cả", type="primary"): 
                for _, row in edited_df_cart.iterrows():
                    service.add_transaction(row["Mã HH"],row["Tên HH"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                    service.update_stock(row["Mã HH"], row["Số lượng"], row["Loại"])
                st.session_state.cart = []
                st.cache_data.clear()
                st.success("🎉 Giao dịch thành công!")
                st.rerun()

elif menu == "Báo cáo tồn kho": show_report()

elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    history = get_cached_history(service)

    if history:
        df = pd.DataFrame(
            history,
            columns=[
                "Ngày",
                "Mã",
                "Tên Hàng Hóa",
                "Loại",
                "Số Lượng",
                "Ghi Chú"
            ]
        )

        df["Số Lượng"] = pd.to_numeric(df["Số Lượng"], errors="coerce").fillna(0)

        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1, minWidth=100)
        
        gb.configure_column("Mã", minWidth=90, maxWidth=130, cellStyle={'textAlign': 'center'})
        gb.configure_column("Tên Hàng Hóa", minWidth=220, cellStyle={'textAlign': 'left'})
        gb.configure_column("Loại", minWidth=90, maxWidth=120, cellStyle={'textAlign': 'center'})
        gb.configure_column(
            "Số Lượng", width=60, suppressSizeToFit=True, type=["numericColumn"], 
            valueFormatter="Number(x).toLocaleString('en-US')", cellStyle={'textAlign': 'right'}
        )
        gb.configure_column("Ghi Chú", minWidth=200, cellStyle={'textAlign': 'left'})

        go = gb.build()
        
        AgGrid(
            df,
            gridOptions=go,
            fit_columns_on_grid_load=True, 
            theme='streamlit',
            height=650 
        )