import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from services.data_service import DataService
from views.report_view_streamlit import show_report

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Quản Lý Kho Hàng", layout="wide")

# CSS TỐI ƯU - ĐẢM BẢO HIỂN THỊ MÀU NÚT 100%
st.markdown("""
    <style>
    /* Nút Thêm vào lưới - Màu Xanh Dương */
    .blue-btn button {
        background-color: #007BFF !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        height: 3em !important;
        width: 100%;
    }
    /* Nút Xác nhận - Màu Xanh Lá */
    .green-btn button {
        background-color: #28a745 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        height: 3em !important;
        width: 100%;
    }
    /* Hiệu ứng khi nhấn */
    .blue-btn button:active, .green-btn button:active {
        transform: scale(0.98);
    }
    </style>
""", unsafe_allow_html=True)

st.title("📦 QL Kho")
service = DataService(mode="ONLINE")

# 2. CACHE DỮ LIỆU
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc): return _svc.get_products()
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc): return _svc.get_history()
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_map(_svc): return _svc.get_product_map()
def clear_all(): st.cache_data.clear()

# 3. ĐIỀU HƯỚNG
menu = st.sidebar.selectbox("Menu", ["Danh mục hàng hóa", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 2: NHẬP/XUẤT ---
if menu == "Nhập/Xuất":
    # Tiêu đề nhỏ bằng 1/2
    st.markdown("#### Nhập/Xuất")
    
    if 'cart' not in st.session_state: st.session_state.cart = []
    
    prod_map = get_cached_map(service)
    # Tối ưu mobile: Placeholder và định dạng số 2,500
    options = {f"{c} - {i['name']} (Tồn: {i['stock']:,.0f} {i['unit']})": c for c, i in prod_map.items()}
    
    col1, col2, col3, col4 = st.columns([2, 1, 1, 2])
    sel = col1.selectbox("Chọn hàng", list(options.keys()), placeholder="Gõ để tìm mã hoặc tên...")
    
    # Sửa lỗi Warning: ép kiểu int ngay tại input
    qty_input = col2.number_input("Số lượng", min_value=1, step=1, value=1)
    qty = int(qty_input)
    
    typ = col3.radio("Loại", ["Nhập", "Xuất"], horizontal=True)
    note = col4.text_input("Ghi chú")
    
    # NÚT THÊM - Bọc trong class blue-btn
    st.markdown('<div class="blue-btn">', unsafe_allow_html=True)
    if st.button("➕ Thêm vào lưới"):
        code = options[sel]
        st.session_state.cart.append({
            "Mã": code, 
            "Loại": typ, 
            "Đvt": prod_map[code]['unit'], 
            "Số lượng": qty, 
            "Ghi chú": note
        })
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.cart:
        st.subheader("📋 Lưới chờ")
        df_cart = pd.DataFrame(st.session_state.cart)
        
        # Chỉ cho sửa Số lượng và Ghi chú
        edited_df = st.data_editor(
            df_cart, use_container_width=True, hide_index=True,
            column_config={
                "Mã": st.column_config.TextColumn(disabled=True),
                "Loại": st.column_config.TextColumn(disabled=True),
                "Đvt": st.column_config.TextColumn(disabled=True),
                "Số lượng": st.column_config.NumberColumn(format="%d", min_value=1)
            }
        )
        
        c1, c2 = st.columns(2)
        
        # NÚT XÁC NHẬN - Bọc trong class green-btn
        with c1:
            st.markdown('<div class="green-btn">', unsafe_allow_html=True)
            if st.button("✅ Xác nhận tất cả"):
                can_proceed = True
                for _, row in edited_df.iterrows():
                    if row["Loại"] == "Xuất":
                        stock = prod_map.get(row["Mã"], {}).get('stock', 0)
                        if row["Số lượng"] > stock:
                            st.error(f"❌ Mã {row['Mã']} không đủ tồn kho (Hiện có: {stock})")
                            can_proceed = False
                
                if can_proceed:
                    for _, row in edited_df.iterrows():
                        service.add_transaction(row["Mã"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                        service.update_stock(row["Mã"], row["Số lượng"], row["Loại"])
                    st.session_state.cart = []
                    clear_all(); st.success("Thành công!"); st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            if st.button("🗑️ Hủy lưới"):
                st.session_state.cart = []; st.rerun()

elif menu == "Báo cáo tồn kho":
    show_report()

elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    hist = get_cached_history(service)
    if hist:
        df_h = pd.DataFrame(hist, columns=["Ngày", "Mã", "Loại", "SL", "Ghi chú"])
        st.dataframe(df_h, use_container_width=True)