import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from controllers.transaction_controller import TransactionController
# Import DataService thay cho ProductController
from services.data_service import DataService 
from datetime import date  
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from openpyxl.styles import PatternFill, Font

def export_to_excel(df):
    # 1. BỎ CỘT AUTO_UNIQUE_ID (nếu có) VÀ CHỈ LẤY CÁC CỘT CẦN THIẾT
    expected_cols = ["Nhóm", "Mã HH", "Tên hàng hóa", "Đvt", "Mức tối thiểu", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]
    available_cols = [col for col in expected_cols if col in df.columns]
    df_export = df[available_cols].copy()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Data')
        ws = writer.sheets['Data']
        ws.freeze_panes = 'A2' # Đóng băng dòng tiêu đề
        
        # =========================================================
        # [TÍNH NĂNG MỚI]: BẬT BỘ LỌC (FILTER) CHO TOÀN BỘ BẢNG
        # =========================================================
        ws.auto_filter.ref = ws.dimensions
        
        # Tự động điều chỉnh độ rộng cột
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 15
            
        # =========================================================
        # 2. TÔ MÀU ĐỎ CHO CÁC MẶT HÀNG CHẠM CẢNH BÁO TỒN KHO
        # =========================================================
        try:
            # Tìm vị trí chính xác của cột "Mức tối thiểu" và "Tồn Cuối"
            min_stock_col_idx = available_cols.index("Mức tối thiểu")
            end_stock_col_idx = available_cols.index("Tồn Cuối")
            
            # Định nghĩa bộ màu chuẩn giống giao diện ứng dụng (Nền đỏ nhạt, Chữ đỏ đậm)
            red_fill = PatternFill(start_color='FFE6E6', end_color='FFE6E6', fill_type='solid')
            red_font = Font(color='D32F2F', bold=True)

            # Quét qua từng dòng trong Excel (bắt đầu từ dòng 2 vì dòng 1 là tiêu đề)
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                # Lấy giá trị của ô Tồn Cuối và Mức tối thiểu tại dòng hiện tại
                min_val = row[min_stock_col_idx].value
                end_val = row[end_stock_col_idx].value
                
                try:
                    # Chuyển sang dạng số để so sánh an toàn
                    v_min = float(min_val) if min_val is not None else 0.0
                    v_end = float(end_val) if end_val is not None else 0.0
                    
                    # Nếu thỏa mãn điều kiện cảnh báo -> Tô màu nguyên cả dòng
                    if v_end <= v_min and v_min > 0:
                        for cell in row:
                            cell.fill = red_fill
                            cell.font = red_font
                except (ValueError, TypeError):
                    continue # Bỏ qua nếu dữ liệu ô bị lỗi (không phải số)
                    
        except ValueError:
            pass # Nếu không tìm thấy cột thì bỏ qua bước tô màu

    return buffer.getvalue()

def show_report():
    st.subheader("Báo cáo tồn kho")
    
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
            margin-top: -1.5rem !important; /* Lực hút đẩy dòng này lên sát tiêu đề */
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
        with st.spinner('Đang kết nối và xử lý dữ liệu...'):
            
            # --- THAY THẾ Ở ĐÂY: GỌI DATASERVICE ---
            service = DataService()
            t_controller = TransactionController()
            
            # Lấy danh mục hàng hóa (trả về danh sách mảng thuần túy)
            products = service.get_products()
            
            # Thêm .copy() để cô lập hoàn toàn bảng dữ liệu, tránh lem cột sang tab Lịch Sử
            df_h = t_controller.get_transaction_history().copy() 
            
            if not products or df_h is None or df_h.empty:
                st.warning("Không có dữ liệu giao dịch hoặc hàng hóa!")
                return

            # --- CHUẨN HÓA DỮ LIỆU ĐỂ TÍNH TOÁN ---
            try:
                df_h['date'] = pd.to_datetime(df_h['Ngày'], dayfirst=True, format='mixed', errors='coerce')
                df_h['product_id'] = df_h['Mã HH'].astype(str).str.strip().str.upper()
                df_h['type'] = df_h['Loại'].astype(str).str.strip()
                df_h['qty'] = pd.to_numeric(df_h['Số Lượng'], errors='coerce').fillna(0)
            except KeyError as e:
                st.error(f"Lỗi cấu trúc cột trong Google Sheets: Thiếu cột {e}")
                return

            # Xác định các mốc biên ngày lọc công thức
            start = pd.to_datetime(start_date).normalize()
            end = pd.to_datetime(end_date).replace(hour=23, minute=59, second=59)
            
            # --- THUẬT TOÁN TÍNH TOÁN TRUY NGƯỢC THỜI GIAN ---
            # 1. Gom toàn bộ giao dịch phát sinh từ NGÀY BẮT ĐẦU cho tới thời điểm HIỆN TẠI
            df_from_start = df_h[df_h['date'] >= start]
            if not df_from_start.empty:
                pivot_from_start = df_from_start.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)
                if 'Nhập' not in pivot_from_start.columns: pivot_from_start['Nhập'] = 0
                if 'Xuất' not in pivot_from_start.columns: pivot_from_start['Xuất'] = 0
            else:
                pivot_from_start = pd.DataFrame(columns=['Nhập', 'Xuất'])
                
            # 2. Gom lượng giao dịch phát sinh giới hạn TRONG KỲ lọc (Từ ngày -> Đến ngày)
            df_period = df_h[(df_h['date'] >= start) & (df_h['date'] <= end)]
            if not df_period.empty:
                pivot_period = df_period.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)
                if 'Nhập' not in pivot_period.columns: pivot_period['Nhập'] = 0
                if 'Xuất' not in pivot_period.columns: pivot_period['Xuất'] = 0
            else:
                pivot_period = pd.DataFrame(columns=['Nhập', 'Xuất'])
            
            # =========================================================
            # --- KHỞI TẠO KHUNG BÁO CÁO TỪ SHEET PRODUCTS ---
            # =========================================================
            product_list = []
            for p in products:
                # p là mảng 7 phần tử: [ID, Mã, Tên, Đvt, Tồn, Nhóm, Mức tối thiểu]
                p_id = p[0]
                p_code = str(p[1]).strip()
                p_name = p[2]
                p_unit = p[3]
                p_stock = p[4]
                p_group = p[5] if len(p) > 5 else ""
                
                # Trích xuất Mức tối thiểu chính xác từ Google Sheets (Cột số 7)
                min_stock = float(p[6]) if len(p) > 6 and str(p[6]).strip() != "" else 0.0
                
                product_list.append([p_id, p_code, p_name, p_unit, p_stock, p_group, min_stock])
            
            # Đưa vào DataFrame và thêm tên cột Nhóm, Mức tối thiểu
            df_products = pd.DataFrame(product_list, columns=["ID", "Mã HH", "Tên hàng hóa", "Đvt", "Tồn Hiện Tại", "Nhóm", "Mức tối thiểu"])
            df_products['Tồn Hiện Tại'] = pd.to_numeric(df_products['Tồn Hiện Tại'], errors='coerce').fillna(0)
            df_products['Mức tối thiểu'] = pd.to_numeric(df_products['Mức tối thiểu'], errors='coerce').fillna(0)
            df_products['Mã HH'] = df_products['Mã HH'].astype(str).str.strip().str.upper()
            
            # Kết nối dữ liệu lũy kế từ ngày lọc đến nay để làm phép tính trừ ngược
            df_report = df_products.merge(pivot_from_start[['Nhập', 'Xuất']], left_on='Mã HH', right_index=True, how='left').fillna(0)
            df_report.rename(columns={'Nhập': 'Nhập_Lũy_Kế', 'Xuất': 'Xuất_Lũy_Kế'}, inplace=True)
            
            # Kết nối dữ liệu thực tế phát sinh trong kỳ
            df_report = df_report.merge(pivot_period[['Nhập', 'Xuất']], left_on='Mã HH', right_index=True, how='left').fillna(0)
            
            # Thực thi thuật toán: 
            df_report['Tồn Đầu'] = df_report['Tồn Hiện Tại'] - df_report['Nhập_Lũy_Kế'] + df_report['Xuất_Lũy_Kế']
            df_report['Tồn Cuối'] = df_report['Tồn Đầu'] + df_report['Nhập'] - df_report['Xuất']
            
            # Định hình lại các cột hiển thị: Kéo cột Nhóm lên đứng đầu tiên và bổ sung Mức tối thiểu
            df_report = df_report[["Nhóm", "Mã HH", "Tên hàng hóa", "Đvt", "Mức tối thiểu", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]]
            
            # =================================================================
            # --- [NÂNG CẤP]: HIỂN THỊ CẢNH BÁO HÀNG SẮP HẾT (DẠNG THU GỌN) ---
            # =================================================================
            st.markdown("---")
            st.subheader("🚨 Cảnh báo mức tồn kho")
            
            # Lọc ra các mặt hàng có Tồn Cuối <= Mức tối thiểu (loại trừ các hàng có Mức tối thiểu = 0)
            df_canh_bao = df_report[(df_report["Tồn Cuối"] <= df_report["Mức tối thiểu"]) & (df_report["Mức tối thiểu"] > 0)]
            
            if not df_canh_bao.empty:
                # Vẫn giữ nguyên dòng báo đỏ để thu hút sự chú ý
                st.error(f"⚠️ Chú ý: Đang có **{len(df_canh_bao)}** mặt hàng chạm mức cảnh báo!")
                
                # Bọc bảng dữ liệu vào trong một expander (khung ẩn/hiện)
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
            # =================================================================

            # --- PHẦN KHỞI TẠO BẢNG AGGRID ---
            gb = GridOptionsBuilder.from_dataframe(df_report)
            gb.configure_default_column(sortable=True, filter=True, resizable=True, flex=1, minWidth=100)
            
            # Kích hoạt gom nhóm theo cột Nhóm
            gb.configure_column("Nhóm", rowGroup=True, hide=True)
            
            gb.configure_column("Mã HH", minWidth=60, maxWidth=120, cellStyle={'textAlign': 'center'})
            gb.configure_column("Tên hàng hóa", minWidth=150, cellStyle={'textAlign': 'left'})
            gb.configure_column("Đvt", minWidth=60, maxWidth=100, cellStyle={'textAlign': 'center'})

            # Cấu hình định dạng số cho tất cả các cột tính toán (Bao gồm cả Mức tối thiểu)
            for col_name in ["Mức tối thiểu", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"]:
                gb.configure_column(
                    col_name,
                    minWidth=90, maxWidth=130,
                    type=["numericColumn"],
                    filter='agNumberColumnFilter',
                    valueFormatter="Number(x).toLocaleString('en-US')",
                    cellStyle={'textAlign': 'right'})
            
            # =================================================================
            # --- THÊM TÍNH NĂNG TÔ MÀU DÒNG CẢNH BÁO BẰNG JSCODE ---
            # =================================================================
            row_style_jscode = JsCode("""
            function(params) {
                // Kiểm tra nếu có dữ liệu và Tồn Cuối <= Mức tối thiểu (Bỏ qua hàng nhóm và Hàng không có định mức)
                if (params.data && params.data['Tồn Cuối'] <= params.data['Mức tối thiểu'] && params.data['Mức tối thiểu'] > 0) {
                    return {
                        'backgroundColor': '#ffe6e6', /* Màu nền đỏ nhạt */
                        'color': '#d32f2f',           /* Chữ màu đỏ đậm */
                        'fontWeight': 'bold'
                    };
                }
                return null;
            }
            """)
            gb.configure_grid_options(getRowStyle=row_style_jscode)
            
            go = gb.build() 
            
            # --- HIỂN THỊ AGGRID ---
            AgGrid(
                df_report,
                gridOptions=go,
                fit_columns_on_grid_load=True,
                theme='streamlit',
                allow_unsafe_jscode=True, # Bắt buộc phải có dòng này để kích hoạt JsCode tô màu
                height=650)
            
            # --- CÁC NÚT BẤM VÀ XUẤT EXCEL ---
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Xuất báo cáo ra Excel (.xlsx)",
                data=export_to_excel(df_report),
                file_name="BaoCaoTonKho.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")