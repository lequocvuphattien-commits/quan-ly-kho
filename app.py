import streamlit as st
import pandas as pd
from services.data_service import DataService
from views.report_view_streamlit import show_report

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Quản Lý Kho Hàng", layout="wide", initial_sidebar_state="auto")    
st.markdown("""
    <style>
    /* Loại bỏ khoảng trắng phía trên tiêu đề chính */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0rem !important;
    }
    /* Loại bỏ khoảng cách thừa phía trên cùng của trang */
    #root > div:nth-child(1) > div > div > div > div > section > div {
        padding-top: 12px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📦 QL Kho")
service = DataService(mode="ONLINE")

# 3. CACHE DỮ LIỆU
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_products(_svc): return _svc.get_products()
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_history(_svc): return _svc.get_history()
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_map(_svc): return _svc.get_product_map()
def clear_all(): st.cache_data.clear()

# 4. ĐIỀU HƯỚNG MENU
menu = st.sidebar.selectbox("Menu", ["Danh mục HH", "Nhập/Xuất", "Báo cáo tồn kho", "Lịch sử giao dịch"])

# --- TAB 1: DANH MỤC ---
if menu == "Danh mục HH":
    st.header("Danh mục HH")
    products = get_cached_products(service)
    if products:
        df = pd.DataFrame(products, columns=["ID", "Mã", "Tên", "Đvt", "Tồn"])
        df["Tồn"] = pd.to_numeric(df["Tồn"], errors="coerce").fillna(0)
        st.dataframe(df[["Mã", "Tên", "Đvt", "Tồn"]], use_container_width=True, hide_index=True)

# --- TAB 2: NHẬP/XUẤT (TỐI ƯU MOBILE & DATA EDITOR) ---
elif menu == "Nhập/Xuất":
    st.markdown("#### Nhập/Xuất")
    if 'cart' not in st.session_state: st.session_state.cart = []
    
    prod_map = get_cached_map(service)
    options = {f"{c} - {i['name']} (Tồn: {i['stock']:,.0f} {i['unit']})": c for c, i in prod_map.items()}
    
    col1, col2, col3, col4 = st.columns([2, 1, 1, 2])
    sel = col1.selectbox("Chọn hàng", list(options.keys()), placeholder="Gõ tìm kiếm...")
    qty = int(col2.number_input("Số lượng", min_value=1, step=1, value=1))
    typ = col3.radio("Loại", ["Nhập", "Xuất"], horizontal=True)
    note = col4.text_input("Ghi chú")
    
    # Nút Thêm (Blue)
    st.markdown('<div class="blue-btn">', unsafe_allow_html=True)
    if st.button("➕ Thêm vào lưới"):
        code = options[sel]
        st.session_state.cart.append({"Mã": code, "Loại": typ, "Đvt": prod_map[code]['unit'], "Số lượng": qty, "Ghi chú": note})
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.cart:
        st.subheader("📋 Lưới chờ (Click ô để sửa)")
        df_cart = pd.DataFrame(st.session_state.cart)
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
        # Nút Xác nhận (Green)
        with c1:
            st.markdown('<div class="green-btn">', unsafe_allow_html=True)
            if st.button("✅ Xác nhận tất cả"):
                for _, row in edited_df.iterrows():
                    if row["Loại"] == "Xuất" and row["Số lượng"] > prod_map.get(row["Mã"], {}).get('stock', 0):
                        st.error(f"❌ Mã {row['Mã']} không đủ tồn kho!"); st.stop()
                for _, row in edited_df.iterrows():
                    service.add_transaction(row["Mã"], row["Số lượng"], row["Loại"], row["Ghi chú"])
                    service.update_stock(row["Mã"], row["Số lượng"], row["Loại"])
                st.session_state.cart = []
                clear_all(); st.success("Hoàn tất!"); st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        if c2.button("🗑️ Hủy"):
            st.session_state.cart = []; st.rerun()

# --- CÁC TAB KHÁC ---
elif menu == "Báo cáo tồn kho":
    show_report()
elif menu == "Lịch sử giao dịch":
    st.header("Lịch sử giao dịch")
    hist = get_cached_history(service)
    if hist:
        df_h = pd.DataFrame(hist, columns=["Ngày", "Mã", "Loại", "SL", "Ghi chú"])
        st.dataframe(df_h, use_container_width=True)