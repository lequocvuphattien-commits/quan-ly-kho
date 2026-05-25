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

    # =========================================================
    # --- [MỚI] CSS ÉP BUỘC NẰM NGANG TRÊN MÀN HÌNH ĐIỆN THOẠI ---
    # =========================================================
    st.markdown("""
        <style>
        /* Ép các cột không được bẻ dòng trên điện thoại */
        [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            flex-direction: row !important;
            gap: 0.5rem !important; /* Khoảng cách giữa các ô */
        }
        /* Cho phép các cột tự động co nhỏ lại cho vừa màn hình điện thoại */
        [data-testid="stHorizontalBlock"] > div {
            min-width: 0px !important; 
        }
        /* Đẩy nút "Báo cáo" thụt xuống ngang hàng với ô nhập ngày */
        [data-testid="stHorizontalBlock"] > div:nth-child(3) {
            padding-top: 1.75rem !important; 
        }
        /* Làm nhỏ chữ ở tiêu đề ô ngày một chút để tránh bị tràn trên màn hình nhỏ */
        [data-testid="stHorizontalBlock"] label {
            font-size: -1rem !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Chia cột với tỉ lệ: Từ ngày (3 phần), Đến ngày (3 phần), Nút bấm (2 phần)
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
            all_history = t_controller.get_transaction_history()
            
            if not products:
                st.warning("Không tìm thấy hàng hóa!")
                return
            if not all_history:
                st.info("Chưa có giao dịch.")
                return

            processed_history = []
            for row in all_history:
                row_copy = list(row)
                while len(row_copy) < 7: 
                    row_copy.append("") 
                processed_history.append(row_copy)

            df_h = pd.DataFrame(processed_history, columns=["date", "product_id", "product_name", "type", "qty", "note", "voucher"])
            
            if not df_h.empty and str(df_h.iloc[0]['date']).strip() == 'Ngày':
                df_h = df_h.iloc[1:].copy()

            df_h['date'] = pd.to_datetime(df_h['date'], errors='coerce').dt.normalize()
            df_h['qty'] = pd.to_numeric(df_h['qty'], errors='coerce').fillna(0)
            df_h['product_id'] = df_h['product_id'].astype(str).str.strip().str.upper()
            df_h['type'] = df_h['type'].astype(str).str.strip().str.capitalize()
            
            df_products = pd.DataFrame(
                [[p.code, p.name, p.unit] for p in products], 
                columns=["code", "name", "unit"]
            )
            df_products['code'] = df_products['code'].astype(str).str.strip().str.upper()
            
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)
            
            df_past = df_h[df_h['date'] < start]
            df_period = df_h[(df_h['date'] >= start) & (df_h['date'] <= end)]
            
            def get_stats(df):
                if df.empty:
                    return pd.DataFrame(columns=['Nhập', 'Xuất'])
                pivot = df.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)
                if 'Nhập' not in pivot.columns: pivot['Nhập'] = 0
                if 'Xuất' not in pivot.columns: pivot['Xuất'] = 0
                return pivot[['Nhập', 'Xuất']]

            past_stats = get_stats(df_past)
            period_stats = get_stats(df_period)
            
            past_stats['ton_dau'] = past_stats['Nhập'] - past_stats['Xuất']
            
            df_report = df_products.merge(past_stats[['ton_dau']], left_on='code', right_index=True, how='left').fillna(0)
            df_report = df_report.merge(period_stats, left_on='code', right_index=True, how='left').fillna(0)
            
            df_report['Tồn Cuối'] = df_report['ton_dau'] + df_report['Nhập'] - df_report['Xuất']
            df_report.columns = ["Mã HH", "Tên hàng hóa", "Đvt", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]
            
            gb = GridOptionsBuilder.from_dataframe(df_report)
            gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1, minWidth=100)
            
            gb.configure_column("Mã HH", minWidth=90, maxWidth=120, cellStyle={'textAlign': 'center'})
            gb.configure_column("Tên hàng hóa", minWidth=150, cellStyle={'textAlign': 'left'})
            gb.configure_column("Đvt", minWidth=80, maxWidth=100, cellStyle={'textAlign': 'center'})

            for col_name in ["Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]:
                gb.configure_column(
                    col_name,
                    minWidth=90, maxWidth=130,
                    type=["numericColumn"],
                    filter='agNumberColumnFilter',
                    valueFormatter="Number(x).toLocaleString('en-US')",
                    cellStyle={'textAlign': 'right'}
                )
            
            go = gb.build()
            
            AgGrid(
                df_report,
                gridOptions=go,
                fit_columns_on_grid_load=True,
                theme='streamlit',
                height=650
            )
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Xuất báo cáo ra Excel (.xlsx)",
                data=export_to_excel(df_report),
                file_name="BaoCaoTonKho.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )