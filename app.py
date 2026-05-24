import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode

# Cấu hình trang (Luôn để đầu tiên)
st.set_page_config(page_title="Quản Lý Kho", layout="wide")

# CSS tinh chỉnh màu sắc nút bấm và giao diện
st.markdown("""
    <style>
    /* Làm nổi bật nút Xác nhận tất cả (Xanh lá) */
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
    }
    /* Chỉnh sửa khoảng cách cho gọn gàng */
    .block-container { padding-top: 1rem !important; }
    </style>
""", unsafe_allow_html=True)

# Khởi tạo dịch vụ
service = DataService(mode="ONLINE")

# --- BỘ NHỚ ĐỆM (CACHE) ---
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc):
    return _svc.get_products()

@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc):
    return _svc.get_history()

st.title("📦 Quản lý kho")

menu = st.sidebar.selectbox("Menu", ["Danh mục hàng", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 1: DANH MỤC HÀNG HÓA ---
if menu == "Danh mục hàng":
    st.header("📋 Danh mục hàng")
    products = get_cached_products(service)
    
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)

        # --- TẠO LƯỚI AG-GRID (GIAO DIỆN TỐI GIẢN CHỈ CÓ Ô NHẬP LIỆU) ---
        gb = GridOptionsBuilder.from_dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]])
        
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
    
    if 'cart' not in st.session_state: 
        st.session_state.cart = []
    
    products = get_cached_products(service)
    
    if products:
        p_dict = {f"{p[1]} - {p[2]}": {"Mã": p[1], "Tên": p[2], "Đvt": p[3], "Tồn": p[4]} for p in products}
        
        with st.container(border=True):
            trans_type = st.radio("Loại giao dịch", ["Nhập", "Xuất"], horizontal=True)
            
            selected = st.selectbox(
                "Chọn hàng hóa", 
                options=list(p_dict.keys()), 
                index=None, 
                placeholder="🔍 Gõ tìm kiếm mã hoặc tên hàng..."
            )
            
            c1, c2, c3, c4 = st.columns([1.2, 1.5, 0.8, 1.5])
            with c1:
                qty = st.number_input("Số lượng", min_value=1.0, value=None, step=1.0, format="%.0f", placeholder="Nhập số...")
            with c2:
                note = st.text_input("Ghi chú", placeholder="Nhập ghi chú...")
            with c3:
                if selected:
                    current_stock = float(p_dict[selected]['Tồn'])
                    unit = p_dict[selected]['Đvt']
                    st.markdown(f"<div style='padding-top: 30px; font-size: 18px; font-weight: bold; color: #28a745;'>Tồn: {current_stock:,.0f} {unit}</div>", unsafe_allow_html=True)
                else:
                    st.markdown("<div style='padding-top: 30px; font-size: 18px;'>Tồn: --</div>", unsafe_allow_html=True)
            with c4:
                st.empty()

            if st.button("➕ Thêm vào lưới"):
                if not selected or not qty:
                    st.warning("⚠️ Vui lòng chọn hàng hóa và nhập số lượng!")
                else:
                    prod_data = p_dict[selected]
                    
                    if trans_type == "Xuất":
                        current_stock = float(prod_data["Tồn"])
                        out_in_cart = sum(item["Số lượng"] for item in st.session_state.cart if item["Mã HH"] == prod_data["Mã"] and item["Loại"] == "Xuất")
                        if qty + out_in_cart > current_stock:
                            st.error(f"❌ Không đủ tồn kho!")
                            st.stop()

                    st.session_state.cart.append({
                        "Mã HH": prod_data["Mã"],
                        "Tên HH": prod_data["Tên"], # Đã thêm cột Tên hàng
                        "Đvt": prod_data["Đvt"],
                        "Số lượng": float(qty),
                        "Ghi chú": note if note else "",
                        "Loại": trans_type
                    })
                    st.rerun() 

        # --- 2. HIỂN THỊ LƯỚI CHỜ ---
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
                hide_index=True,
                use_container_width=True,
                key="cart_editor"
            )
            
            col_x, col_y = st.columns([1, 5])
            with col_x:
                if st.button("✅ Xác nhận tất cả", type="primary"): 
                    for _, row in edited_df.iterrows():
                        service.add_transaction(row["Mã HH"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                        service.update_stock(row["Mã HH"], row["Số lượng"], row["Loại"])
                    
                    st.session_state.cart = []
                    st.cache_data.clear()
                    st.success("🎉 Giao dịch thành công!")
                    st.rerun()
            with col_y:
                if st.button("🗑️ Hủy lưới"):
                    st.session_state.cart = []
                    st.rerun()

# --- TAB 3: BÁO CÁO TỒN KHO ---
elif menu == "Báo cáo tồn kho":
    show_report()

# --- TAB 4: LỊCH SỬ ---
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    history = get_cached_history(service)
    if history:
        st.dataframe(pd.DataFrame(history, columns=["Ngày", "Mã HH", "Loại", "Số Lượng", "Ghi Chú"]), width='stretch', hide_index=True)
