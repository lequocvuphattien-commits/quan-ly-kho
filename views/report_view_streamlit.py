import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from controllers.transaction_controller import TransactionController
from controllers.product_controller import ProductController
from datetime import date  
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode

def export_to_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        ws.freeze_panes = 'A2'
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 15
    return buffer.getvalue()

def show_report():
    st.subheader("Báo cáo tồn kho")
    
    p_controller = ProductController()
    t_controller = TransactionController()
    
    DEFAULT_START_DATE = date(2026, 1, 1)
    
    if "clicked_report_filter" not in st.session_state:
        st.session_state.clicked_report_filter = False

    # =================================================================
    # --- [CẢI TIẾN CSS]: ÉP HÀNG NGANG VÀ KÉO SÁT TIÊU ĐỀ TRÊN ĐIỆN THOẠI ---
    # =================================================================
    st.markdown("""
        <style>
        /* 1. Kéo tiêu đề h3 dịch xuống hoặc thu nhỏ khoảng trống bên dưới nó */
        h3 {
            margin-bottom: -0.5rem !important;
            padding-bottom: 0rem !important;
        }
        
        /* 2. Ép 3 thành phần nằm ngang và kéo mạnh lên sát tiêu đề h3 */
        [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            flex-direction: row !important;
            gap: 0.4rem !important; /* Thu hẹp khoảng cách giữa các ô */
            margin-top: -1.5rem !important; /* [QUAN TRỌNG] Lực hút đẩy dòng này lên sát tiêu đề */
        }
        
        /* Cho phép các ô tự động co nhỏ cho vừa màn hình điện thoại */
        [data-testid="stHorizontalBlock"] > div {
            min-width: 0px !important; 
        }
        
        /* Đẩy nút "Báo cáo" thụt xuống để thẳng hàng khít với ô nhập ngày */
        [data-testid="stHorizontalBlock"] > div:nth-child(3) {
            padding-top: 1.75rem !important; 
        }
        
        /* Làm nhỏ chữ tiêu đề "Từ ngày", "Đến ngày" để giao diện thanh thoát */
        [data-testid="stHorizontalBlock"] label {
            font-size: 0.8rem !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Chia cột tỉ lệ phù hợp cho màn hình dọc điện thoại
    col1, col2, col3 = st.columns([3, 3, 2.5])
    
    with col1:
        start_date = st.date_input("Từ ngày", value=DEFAULT_START_DATE)
    
    with col2:
        end_date = st.date_input("Đến ngày")
        
    with col3:
        if st.button("Báo cáo", type="primary", use_container_width=True):
            st.session_state.clicked_report_filter = True

    # --- NẾU ĐÃ BẤM NÚT LỌC, TIẾN HÀNH XỬ LÝ VÀ HIỂN THỊ DỮ LIỆU ---
    if st.session_state.clicked_report_filter:
        with st.spinner('Đang xử lý dữ liệu...'):
            products = p_controller.get_all_products()
            df_h = t_controller.get_transaction_history() # Trả về DataFrame
            
            # Kiểm tra dữ liệu an toàn (Dùng .empty cho DataFrame)
            if not products or df_h is None or df_h.empty:
                st.warning("Không có dữ liệu giao dịch hoặc hàng hóa!")
                return

            # --- CHUẨN HÓA DỮ LIỆU ĐỂ TÍNH TOÁN ---
            try:
                # Đảm bảo các cột được ép kiểu đúng
                df_h['date'] = pd.to_datetime(df_h['Ngày'], errors='coerce')
                df_h['product_id'] = df_h['Mã HH'].astype(str).str.strip().str.upper()
                df_h['type'] = df_h['Loại'].astype(str).str.strip()
                df_h['qty'] = pd.to_numeric(df_h['Số Lượng'], errors='coerce').fillna(0)
            except KeyError as e:
                st.error(f"Lỗi cấu trúc cột trong Google Sheets: Thiếu cột {e}")
                return

            # --- TÍNH TOÁN TỒN KHO ---
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)
            
            # Lọc dữ liệu dựa trên ngày
            df_past = df_h[df_h['date'] < start]
            df_period = df_h[(df_h['date'] >= start) & (df_h['date'] <= end)]
            
            def get_stats(df):
                if df.empty: return pd.DataFrame(columns=['Nhập', 'Xuất'])
                pivot = df.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)
                # Đảm bảo cột luôn tồn tại
                if 'Nhập' not in pivot.columns: pivot['Nhập'] = 0
                if 'Xuất' not in pivot.columns: pivot['Xuất'] = 0
                return pivot[['Nhập', 'Xuất']]

            past_stats = get_stats(df_past)
            past_stats['ton_dau'] = past_stats['Nhập'] - past_stats['Xuất']
            period_stats = get_stats(df_period)
            
            # --- TẠO BÁO CÁO CUỐI CÙNG ---
            df_products = pd.DataFrame([[p.code, p.name, p.unit] for p in products], 
                                     columns=["Mã HH", "Tên hàng hóa", "Đvt"])
            df_products['Mã HH'] = df_products['Mã HH'].astype(str).str.strip().str.upper()
            
            df_report = df_products.merge(past_stats[['ton_dau']], left_on='Mã HH', right_index=True, how='left').fillna(0)
            df_report = df_report.merge(period_stats, left_on='Mã HH', right_index=True, how='left').fillna(0)
            
            df_report['Tồn Cuối'] = df_report['ton_dau'] + df_report['Nhập'] - df_report['Xuất']
            # Đổi tên cột hiển thị
            df_report.columns = ["Mã HH", "Tên hàng hóa", "Đvt", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]
            
            # --- PHẦN KHỞI TẠO BẢNG AGGRID ---
            # Đảm bảo gb và go được định nghĩa bên trong khối này
            gb = GridOptionsBuilder.from_dataframe(df_report)
            gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1, minWidth=100)
            
            gb.configure_column("Mã HH", minWidth=60, maxWidth=120, cellStyle={'textAlign': 'center'})
            gb.configure_column("Tên hàng hóa", minWidth=150, cellStyle={'textAlign': 'left'})
            gb.configure_column("Đvt", minWidth=60, maxWidth=100, cellStyle={'textAlign': 'center'})

            for col_name in ["Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]:
                gb.configure_column(
                    col_name,
                    minWidth=90, maxWidth=130,
                    type=["numericColumn"],
                    filter='agNumberColumnFilter',
                    valueFormatter="Number(x).toLocaleString('en-US')",
                    cellStyle={'textAlign': 'right'})
            
            go = gb.build() # Biến go được tạo ra ở đây
            # --- HIỂN THỊ AGGRID ---
            AgGrid(
                df_report,
                gridOptions=go,
                fit_columns_on_grid_load=True,
                theme='streamlit',
                height=650)
            
            # --- CÁC NÚT BẤM VÀ XUẤT EXCEL ---
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Xuất báo cáo ra Excel (.xlsx)",
                data=export_to_excel(df_report),
                file_name="BaoCaoTonKho.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")