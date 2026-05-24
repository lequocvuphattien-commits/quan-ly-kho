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

# Cấu hình trang (Luôn để đầu tiên)
st.set_page_config(page_title="Quản Lý Kho", layout="wide")

# CSS tinh chỉnh màu sắc nút bấm và giao diện
st.markdown("""
    <style>
    /* Làm nổi bật nút Xác nhận tất cả (Xanh lá) */
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;}         
    </style>
""", unsafe_allow_html=True)

# Khởi tạo dịch vụ
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

        # Cấu hình AgGrid với tính năng lọc nâng cao và sắp xếp
        gb = GridOptionsBuilder.from_dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]])

         # --- THÊM DÒNG NÀY ĐỂ CỘT ĐVT VỪA VỚI DỮ LIỆU ---
        gb.configure_column("Mã", width=60, suppressSizeToFit=True) # Cột Mã nhỏ gọn, canh giữa
        gb.configure_column("Đvt", width=60, suppressSizeToFit=True, cellStyle={'textAlign': 'center'}) # Cột Đvt nhỏ gọn, canh giữa
        gb.configure_column("Tên", width=200, cellStyle={'textAlign': 'left'}) # Cột Tên rộng hơn, canh trái

        # --- CẬP NHẬT LẠI CỘT TỒN: THÊM ĐỊNH DẠNG DẤU PHẨY ---
        gb.configure_column(
            "Tồn", 
            width=60, 
            suppressSizeToFit=True, 
            type=["numericColumn"], # Khai báo là cột số
            valueFormatter="Number(x).toLocaleString('en-US')", # Format hàng nghìn (3,010)
            cellStyle={'textAlign': 'right'}
        )
        

        gb.configure_default_column(
            sortable=True,
            filter=True,
            floatingFilter=True, 
            resizable=True,
            suppressMenu=True, 
            filterParams={"suppressFilterButton": True} 
        )
        go = gb.build()
        if 'columnDefs' in go:
            for col in go['columnDefs']:
                if col.get('field') == 'Tồn':
                    col['filter'] = 'agNumberColumnFilter'

        grid_response = AgGrid(
            df[["Mã", "Tên", "Đvt", "Tồn"]],
            gridOptions=go,
            fit_columns_on_grid_load=True,
            theme='streamlit',
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED 
        )
        
        filtered_df = pd.DataFrame(grid_response['data'])
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if not filtered_df.empty:
                filtered_df.to_excel(writer, index=False, sheet_name='DanhMuc')
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

# --- TAB 2: NHẬP/XUẤT ---
elif menu == "Nhập/Xuất":
    st.subheader("🔄 Nhập/Xuất kho")
    
    # 1. Đặt radio trước để khởi tạo biến trans_type
    trans_type = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
    
    # 2. Khai báo danh sách kho ngay sau khi biết trans_type
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
                    
        # 3. Đưa phần Thêm kho vào đây (đã có trans_type)
        with st.expander(f"➕ Thêm địa điểm mới: {trans_type}"):
            new_kho = st.text_input("Tên địa điểm mới", placeholder="Ví dụ: Kho ...")
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
            edited_df = st.data_editor(
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
                for _, row in edited_df.iterrows():
                    service.add_transaction(row["Mã HH"],row["Tên HH"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                    service.update_stock(row["Mã HH"], row["Số lượng"], row["Loại"])
                st.session_state.cart = []
                st.cache_data.clear()
                st.success("🎉 Giao dịch thành công!")
                st.rerun()

elif menu == "Báo cáo tồn kho": show_report() # Gọi hàm hiển thị báo cáo tồn kho

elif menu == "Lịch sử giao dịch":

    st.header("Lịch sử giao dịch")

    history = get_cached_history(service)

    if history:

        # Tạo DataFrame
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

        # Ép kiểu số
        df["Số Lượng"] = pd.to_numeric(
            df["Số Lượng"],
            errors="coerce"
        ).fillna(0)

        # Khởi tạo AgGrid
        gb = GridOptionsBuilder.from_dataframe(df)

        # Cấu hình mặc định
        gb.configure_default_column(
            sortable=True,
            filter=True,
            resizable=True,
            flex=1,
            minWidth=100
        )

        # Cột Mã
        gb.configure_column(
            "Mã",
            minWidth=90,
            maxWidth=130,
            cellStyle={'textAlign': 'center'}
        )

        # Cột Tên Hàng Hóa
        gb.configure_column(
            "Tên Hàng Hóa",
            minWidth=220,
            cellStyle={'textAlign': 'left'}
        )

        # Cột Loại
        gb.configure_column(
            "Loại",
            minWidth=90,
            maxWidth=120,
            cellStyle={'textAlign': 'center'}
        )

        # Cột Số Lượng
        gb.configure_column(
            "Số Lượng", 
            width=60, 
            suppressSizeToFit=True, 
            type=["numericColumn"], # Khai báo là cột số
            valueFormatter="Number(x).toLocaleString('en-US')", # Format hàng nghìn (3,010)
            cellStyle={'textAlign': 'right'}
        )

        # Cột Ghi Chú
        gb.configure_column(
            "Ghi Chú",
            minWidth=200,
            cellStyle={'textAlign': 'left'}
        )

        # Build grid
        go = gb.build()

        # Hiển thị bảng
        AgGrid(
            df,
            gridOptions=go,

            # Tự co giãn theo dữ liệu
            fit_columns_on_grid_load=True,

            theme='streamlit',

            # Full width
            use_container_width=True
        )


