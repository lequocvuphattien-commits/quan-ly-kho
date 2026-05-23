import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Quản Lý Kho Hàng", layout="wide")
st.title("📦 QL Kho")

# 2. KHỞI TẠO DỊCH VỤ
service = DataService(mode="ONLINE")

# 3. BỘ NHỚ ĐỆM (CACHE)
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc): return _svc.get_products()
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc): return _svc.get_history()
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_map(_svc): return _svc.get_product_map()

def clear_all_caches(): st.cache_data.clear()

# 4. HÀM XUẤT EXCEL CHUYÊN NGHIỆP
def export_to_excel(df, sheet_name='Data'):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        worksheet.freeze_panes = 'A2'
        max_row, max_col = worksheet.max_row, worksheet.max_column
        worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
        for col in range(1, max_col + 1):
            worksheet.column_dimensions[get_column_letter(col)].width = 15
    return buffer.getvalue()

# 5. MENU ĐIỀU HƯỚNG
menu = st.sidebar.selectbox("Menu", ["Danh mục hàng hóa", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# 6. CÁC TAB CHỨC NĂNG
# --- TAB 1: DANH MỤC ---
if menu == "Danh mục hàng hóa":
    st.header("Danh mục hàng hóa")
    products = get_cached_products(service)
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        st.dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]], use_container_width=True, hide_index=True)
        st.download_button("📥 Xuất danh mục (.xlsx)", export_to_excel(df[["Mã", "Tên", "Đvt", "Tồn"]]), "DanhMuc.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    with st.form("add_form", clear_on_submit=True):
        st.subheader("Thêm hàng hóa mới")
        c, n, u = st.text_input("Mã hàng"), st.text_input("Tên hàng"), st.text_input("Đơn vị tính")
        if st.form_submit_button("Thêm hàng hóa"):
            if not c or not n: st.warning("Nhập đủ Mã và Tên!")
            elif service.check_product_exists(c.upper()): st.error("Mã đã tồn tại!")
            else:
                service.add_product(c, n, u)
                clear_all_caches(); st.success("Đã thêm!"); st.rerun()

# --- TAB 2: NHẬP/XUẤT (TỐI ƯU HÀNG LOẠT) ---
elif menu == "Nhập/Xuất":
    st.markdown("### Nhâp/Xuất")
    if 'cart' not in st.session_state: st.session_state.cart = []
    
    prod_map = get_cached_map(service)
    options = {f"{c} - {i['name']} (Tồn: {i['stock']:,.0f} {i['unit']})": c for c, i in prod_map.items()}
    
    col1, col2, col3, col4 = st.columns([2, 1, 1, 2])
    # Tối ưu cho mobile: thêm placeholder
    sel = col1.selectbox("Chọn hàng", list(options.keys()), placeholder="Gõ tìm kiếm...")
    # Khắc phục cảnh báo: Ép kiểu int() ngay từ đầu
    qty = int(col2.number_input("Số lượng", min_value=1, step=1, format="%d"))
    typ = col3.radio("Loại", ["Nhập", "Xuất"], horizontal=True)
    note = col4.text_input("Ghi chú")
    
    if st.button("➕ Thêm vào lưới"):
        code = options[sel]
        st.session_state.cart.append({"Mã": code, "Loại": typ, "Đvt": prod_map[code]['unit'], "Số lượng": qty, "Ghi chú": note})
        st.rerun()

    if st.session_state.cart:
        st.markdown("📋 Chờ giao dịch")
        df_cart = pd.DataFrame(st.session_state.cart)
        
        # CẤU HÌNH DATA_EDITOR: Chặn sửa Mã, Loại, Đvt
        edited_df = st.data_editor(
            df_cart, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Mã": st.column_config.TextColumn(disabled=True),
                "Loại": st.column_config.TextColumn(disabled=True),
                "Đvt": st.column_config.TextColumn(disabled=True),
                "Số lượng": st.column_config.NumberColumn(format="%d", min_value=1)
            }
        )
        
        # Kiểm tra cảnh báo xuất quá tồn kho sau khi sửa trên lưới
        if st.button("✅ Xác nhận tất cả"):
            can_proceed = True
            for _, row in edited_df.iterrows():
                if row["Loại"] == "Xuất":
                    current_stock = prod_map.get(row["Mã"], {}).get('stock', 0)
                    if row["Số lượng"] > current_stock:
                        st.error(f"❌ Mã {row['Mã']} chỉ còn {current_stock}, không thể xuất {row['Số lượng']}!")
                        can_proceed = False
            
            if can_proceed:
                for _, row in edited_df.iterrows():
                    service.add_transaction(row["Mã"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                    service.update_stock(row["Mã"], row["Số lượng"], row["Loại"])
                st.session_state.cart = []
                clear_all_caches(); st.success("Hoàn tất!"); st.rerun()

# --- TAB 3: BÁO CÁO ---
elif menu == "Báo cáo tồn kho":
    show_report()

# --- TAB 4: LỊCH SỬ ---
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    hist = get_cached_history(service)
    if hist:
        df_h = pd.DataFrame(hist, columns=["Ngày", "Mã", "Loại", "SL", "Ghi chú"])
        st.dataframe(df_h, use_container_width=True)
        st.download_button("📥 Xuất lịch sử (.xlsx)", export_to_excel(df_h), "LichSu.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")