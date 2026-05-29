import streamlit as st
import pandas as pd
import io
import os
from openpyxl.utils import get_column_letter
from controllers.transaction_controller import TransactionController
from services.data_service import DataService 
from datetime import date  
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.drawing.image import Image

# --- [TỐI ƯU TỐC ĐỘ]: LẤY DỮ LIỆU BẰNG CACHE KHÔNG LOAD LẠI API ---
@st.cache_data(ttl=600, show_spinner=False)
def get_report_cached_products():
    return DataService(mode="ONLINE").get_products()

@st.cache_data(ttl=600, show_spinner=False)
def get_report_cached_history():
    return TransactionController().get_transaction_history().copy()

def export_to_excel(df, end_date):
    expected_cols = ["Nhóm", "Mã HH", "Tên hàng hóa", "Đvt", "Mức tối thiểu", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối", "Ghi chú"]
    available_cols = [col for col in expected_cols if col in df.columns]
    df_export = df[available_cols].copy()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Data', startrow=6)
        ws = writer.sheets['Data']
        
        ws.freeze_panes = 'A8' 
        ws.auto_filter.ref = f"A7:{get_column_letter(ws.max_column)}{ws.max_row}"
        
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 15
            
        header_fill = PatternFill(start_color='0070C0', end_color='0070C0', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=7, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws['C1'] = "CÔNG TY TNHH THỦY SẢN PHÁT TIẾN"
        ws['C1'].font = Font(size=14,bold=True)
        ws['C2'] = "Địa chỉ: Lô B3, đường số 2, Cụm CN Mỹ Hiệp, Xã Mỹ Hiệp, Tỉnh Đồng Tháp"
        ws['C2'].font = Font(name="Arial", size=10)
        ws['C3'] = "Số điện thoại: 02778.553.388 - 02773.918.999"
        ws['C3'].font = Font(name="Arial", size=10)

        ws.merge_cells('D4:E4')
        ws['D4'] = "BÁO CÁO TỒN KHO VẬT TƯ"
        ws['D4'].font = Font(bold=True, size=14)
        ws['D4'].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells('D5:E5')
        if end_date:
            ws['D5'] = end_date.strftime("Ngày %d tháng %m năm %Y")
        ws['D5'].font = Font(italic=True)
        ws['D5'].alignment = Alignment(horizontal="center", vertical="center")

        try:
            logo_path = "logo.png"
            if os.path.exists(logo_path):
                img = Image(logo_path)
                img.width = 100
                img.height = 100
                ws.add_image(img, 'A1')
        except Exception as e:
            print(f"Lỗi chèn ảnh Logo: {e}")

        try:
            min_stock_col_idx = available_cols.index("Mức tối thiểu")
            end_stock_col_idx = available_cols.index("Tồn Cuối")
            
            red_fill = PatternFill(start_color='FFE6E6', end_color='FFE6E6', fill_type='solid')
            red_font = Font(color='D32F2F', bold=True)

            for row in ws.iter_rows(min_row=8, max_row=ws.max_row):
                min_val = row[min_stock_col_idx].value
                end_val = row[end_stock_col_idx].value
                
                try:
                    v_min = float(min_val) if min_val is not None else 0.0
                    v_end = float(end_val) if end_val is not None else 0.0
                    
                    if v_end <= v_min and v_min > 0:
                        for cell in row:
                            cell.fill = red_fill
                            cell.font = red_font
                except (ValueError, TypeError):
                    continue 
                    
        except ValueError:
            pass 

    return buffer.getvalue()

def export_history_to_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='LichSu')
        ws = writer.sheets['LichSu']
        header_fill = PatternFill(start_color='0070C0', end_color='0070C0', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(col)].width = 18
    return buffer.getvalue()

def show_report():
    st.subheader("Báo cáo tồn kho")
    
    DEFAULT_START_DATE = date(2026, 1, 1)
    
    if "clicked_report_filter" not in st.session_state:
        st.session_state.clicked_report_filter = False

    st.markdown("""
        <style>
        h3 {
            margin-bottom: -0.5rem !important;
            padding-bottom: 0rem !important;
        }
        [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            flex-direction: row !important;
            gap: 0.4rem !important;
            margin-top: -1.5rem !important; 
        }
        [data-testid="stHorizontalBlock"] > div {
            min-width: 0px !important; 
        }
        [data-testid="stHorizontalBlock"] > div:nth-child(3) {
            padding-top: 1.75rem !important; 
        }
        [data-testid="stHorizontalBlock"] label {
            font-size: 0.8rem !important;
        }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([3, 3, 2.5])
    
    with col1:
        start_date = st.date_input("Từ ngày", value=DEFAULT_START_DATE)
    
    with col2:
        end_date = st.date_input("Đến ngày")
        
    with col3:
        if st.button("Báo cáo", type="primary", use_container_width=True):
            st.session_state.clicked_report_filter = True

    if st.session_state.clicked_report_filter:
        with st.spinner('Đang kết nối và xử lý dữ liệu...'):
            
            # =================================================================
            # --- [TỐI ƯU UX & TỐC ĐỘ]: CACHE UI BÁO CÁO (KHÔNG TÍNH TOÁN LẠI)---
            # =================================================================
            current_cache_key = f"{start_date}_{end_date}"
            
            if st.session_state.get("report_cache_key") != current_cache_key or "report_df" not in st.session_state:
                
                # SỬ DỤNG HÀM CACHE SIÊU TỐC THAY VÌ KHỞI TẠO SERVICE GỌI API LẠI
                products = get_report_cached_products()
                df_h = get_report_cached_history()
                
                if not products or df_h is None or df_h.empty:
                    st.warning("Không có dữ liệu giao dịch hoặc hàng hóa!")
                    return

                try:
                    df_h['date'] = pd.to_datetime(df_h['Ngày'], dayfirst=True, format='mixed', errors='coerce')
                    df_h['Ngày_Kho'] = (df_h['date'] - pd.Timedelta(hours=6)).dt.normalize()
                    df_h['product_id'] = df_h['Mã HH'].astype(str).str.strip().str.upper()
                    df_h['type'] = df_h['Loại'].astype(str).str.strip()
                    df_h['qty'] = pd.to_numeric(df_h['Số Lượng'], errors='coerce').fillna(0)
                except KeyError as e:
                    st.error(f"Lỗi cấu trúc cột trong Google Sheets: Thiếu cột {e}")
                    return

                start = pd.to_datetime(start_date).normalize()
                end = pd.to_datetime(end_date).normalize()
                
                df_from_start = df_h[df_h['Ngày_Kho'] < start]
                if not df_from_start.empty:
                    pivot_from_start = df_from_start.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)
                    if 'Nhập' not in pivot_from_start.columns: pivot_from_start['Nhập'] = 0
                    if 'Xuất' not in pivot_from_start.columns: pivot_from_start['Xuất'] = 0
                else:
                    pivot_from_start = pd.DataFrame(columns=['Nhập', 'Xuất'])
                    
                df_period = df_h[(df_h['Ngày_Kho'] >= start) & (df_h['Ngày_Kho'] <= end)]
                if not df_period.empty:
                    pivot_period = df_period.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)
                    if 'Nhập' not in pivot_period.columns: pivot_period['Nhập'] = 0
                    if 'Xuất' not in pivot_period.columns: pivot_period['Xuất'] = 0
                else:
                    pivot_period = pd.DataFrame(columns=['Nhập', 'Xuất'])

                # --- [TỐI ƯU TỐC ĐỘ]: LƯỢT BỎ FOR LOOP CHẬM CHẠP ---
                product_list = [
                    [
                        p[0], str(p[1]).strip().upper(), p[2], p[3], float(p[4]) if p[4] else 0.0,
                        p[5] if len(p) > 5 else "",
                        float(p[6]) if len(p) > 6 and str(p[6]).strip() != "" else 0.0,
                        str(p[7]).strip() if len(p) > 7 else ""
                    ] for p in products
                ]
                
                df_products = pd.DataFrame(product_list, columns=["ID", "Mã HH", "Tên hàng hóa", "Đvt", "Tồn Hiện Tại", "Nhóm", "Mức tối thiểu", "Ghi chú"])
                
                df_report = df_products.merge(pivot_from_start[['Nhập', 'Xuất']], left_on='Mã HH', right_index=True, how='left').fillna(0)
                df_report.rename(columns={'Nhập': 'Nhập_Lũy_Kế', 'Xuất': 'Xuất_Lũy_Kế'}, inplace=True)
                
                df_report = df_report.merge(pivot_period[['Nhập', 'Xuất']], left_on='Mã HH', right_index=True, how='left').fillna(0)
                
                df_report['Tồn Đầu'] = df_report['Tồn Hiện Tại'] - df_report['Nhập_Lũy_Kế'] + df_report['Xuất_Lũy_Kế']
                df_report['Tồn Cuối'] = df_report['Tồn Đầu'] + df_report['Nhập'] - df_report['Xuất']
                
                df_report = df_report[["Nhóm", "Mã HH", "Tên hàng hóa", "Đvt", "Mức tối thiểu", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối", "Ghi chú"]]
                
                # --- CHUẨN BỊ LƯỚI CHO LẦN ĐẦU ---
                gb = GridOptionsBuilder.from_dataframe(df_report)
                gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1, minWidth=100)
                gb.configure_column("Nhóm", rowGroup=True, hide=True)
                gb.configure_column("Mã HH", minWidth=60, maxWidth=120, cellStyle={'textAlign': 'center'})
                gb.configure_column("Tên hàng hóa", minWidth=150, cellStyle={'textAlign': 'left'})
                gb.configure_column("Đvt", minWidth=60, maxWidth=100, cellStyle={'textAlign': 'center'})
                gb.configure_column("Ghi chú", minWidth=120, cellStyle={'textAlign': 'left'})

                for col_name in ["Mức tối thiểu", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]:
                    gb.configure_column(
                        col_name,
                        minWidth=90, maxWidth=130,
                        type=["numericColumn"],
                        filter='agNumberColumnFilter',
                        valueFormatter="Number(x).toLocaleString('en-US')",
                        cellStyle={'textAlign': 'right'})
                
                row_style_jscode = JsCode("""
                function(params) {
                    if (params.data && params.data['Tồn Cuối'] <= params.data['Mức tối thiểu'] && params.data['Mức tối thiểu'] > 0) {
                        return { 'backgroundColor': '#ffe6e6', 'color': '#d32f2f', 'fontWeight': 'bold' };
                    }
                    return null;
                }
                """)
                gb.configure_grid_options(getRowStyle=row_style_jscode)
                
                # LƯU VÀO CACHE UI (Không bao giờ tính lại trừ khi đổi ngày)
                st.session_state.report_df = df_report
                st.session_state.report_go = gb.build()
                st.session_state.report_cache_key = current_cache_key

            # ==============================================================
            # HIỂN THỊ TRỰC TIẾP TỪ CACHE (CHỚP MẮT 0.1s LÀ XONG)
            # ==============================================================
            df_report_cached = st.session_state.report_df
            go_cached = st.session_state.report_go

            st.markdown("---")
            st.subheader("🚨 Cảnh báo mức tồn kho")
            
            df_canh_bao = df_report_cached[(df_report_cached["Tồn Cuối"] <= df_report_cached["Mức tối thiểu"]) & (df_report_cached["Mức tối thiểu"] > 0)]
            
            if not df_canh_bao.empty:
                st.error(f"⚠️ Chú ý: Đang có **{len(df_canh_bao)}** mặt hàng chạm mức cảnh báo!")
                
                with st.expander("👇 Bấm vào đây để xem chi tiết danh sách hàng hóa", expanded=False):
                    st.dataframe(
                        df_canh_bao[["Mã HH", "Tên hàng hóa", "Đvt", "Tồn Cuối", "Mức tối thiểu"]], 
                        use_container_width=True, 
                        hide_index=True
                    )
            else:
                st.success("✅ Tuyệt vời! Tất cả hàng hóa đều có số lượng tồn kho trên mức an toàn.")
                
            st.markdown("---")
            st.subheader("📦 Chi tiết tồn kho toàn bộ hàng hóa")

            AgGrid(
                df_report_cached,
                gridOptions=go_cached,
                fit_columns_on_grid_load=True,
                theme='streamlit',
                allow_unsafe_jscode=True, 
                height=650)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Xuất báo cáo ra Excel (.xlsx)",
                data=export_to_excel(df_report_cached, end_date), 
                file_name=f"BaoCaoTonKho_{end_date.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
def get_count_import_by_department(service, start_date, end_date):
    history_df = service.get_history()
    if history_df is None or history_df.empty:
        return pd.DataFrame(columns=['Bộ phận', 'Số lần nhập'])

    df = history_df.copy()
    
    df.columns = [str(col).strip() for col in df.columns]
    
    col_bp = 'Bộ phận'
    col_date = 'Ngày'
    col_type = 'Loại'
    
    if col_bp not in df.columns or col_date not in df.columns or col_type not in df.columns:
        return pd.DataFrame(columns=[col_bp, 'Số lần nhập'])

    df_nhap = df[df[col_type].astype(str).str.strip().str.upper() == 'NHẬP'].copy()
    if df_nhap.empty:
        return pd.DataFrame(columns=[col_bp, 'Số lần nhập'])
    
    df_nhap['date_parsed'] = pd.to_datetime(df_nhap[col_date], dayfirst=True, errors='coerce')
    df_nhap['just_date'] = df_nhap['date_parsed'].dt.date
    
    start = pd.to_datetime(start_date).date()
    end = pd.to_datetime(end_date).date()
    
    df_nhap = df_nhap[(df_nhap['just_date'] >= start) & (df_nhap['just_date'] <= end)]
    if df_nhap.empty:
        return pd.DataFrame(columns=[col_bp, 'Số lần nhập'])

    df_nhap[col_bp] = df_nhap[col_bp].astype(str).str.strip()
    df_nhap.loc[df_nhap[col_bp] == "", col_bp] = "Chưa xác định"
    
    df_nhap['Thời_gian_chính_xác'] = df_nhap[col_date].astype(str).str.strip()

    df_unique = df_nhap.drop_duplicates(subset=['Thời_gian_chính_xác', col_bp])
    
    report = df_unique.groupby(col_bp).size().reset_index(name='Số lần nhập')
    report = report.sort_values(by='Số lần nhập', ascending=False).reset_index(drop=True)
    
    return report