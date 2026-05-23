import streamlit as st
import pandas as pd
import io
from openpyxl.utils import get_column_letter
from controllers.transaction_controller import TransactionController
from controllers.product_controller import ProductController

def export_to_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        worksheet = writer.sheets['Data']
        worksheet.freeze_panes = 'A2'
        max_row, max_col = worksheet.max_row, worksheet.max_column
        worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
        for col in range(1, max_col + 1):
            worksheet.column_dimensions[get_column_letter(col)].width = 15
    return buffer.getvalue()

def show_report():
    st.subheader("Báo cáo tồn kho")
    
    p_controller = ProductController()
    t_controller = TransactionController()
    
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Từ ngày")
    end_date = col2.date_input("Đến ngày")
    
    if st.button("Lọc báo cáo", type="primary"):
        with st.spinner('Đang xử lý dữ liệu...'):
            products = p_controller.get_all_products()
            all_history = t_controller.get_transaction_history()
            
            if not products:
                st.warning("Không tìm thấy hàng hóa!")
                return
            if not all_history:
                st.info("Chưa có giao dịch nào.")
                return

            df_h = pd.DataFrame(all_history, columns=["date", "product_id", "type", "qty", "note"])
            df_h['date'] = pd.to_datetime(df_h['date'])
            df_h['qty'] = pd.to_numeric(df_h['qty'], errors='coerce').fillna(0)
            df_h['type'] = df_h['type'].astype(str).str.strip()
            
            start, end = pd.to_datetime(start_date), pd.to_datetime(end_date)
            
            # --- TỐI ƯU HÓA: Tính toán bằng Pivot Table ---
            # Chia dữ liệu thành 2 nhóm: Trước kỳ và Trong kỳ
            df_past = df_h[df_h['date'] < start]
            df_period = df_h[(df_h['date'] >= start) & (df_h['date'] <= end)]
            
            # Tính toán nhanh bằng Pivot
            def get_pivot(df):
                return df.pivot_table(index='product_id', columns='type', values='qty', aggfunc='sum', fill_value=0)

            past_pivot = get_pivot(df_past)
            period_pivot = get_pivot(df_period)
            
            report_data = []
            for p in products:
                pid = str(p.code).strip()
                
                # Tồn đầu = Nhập đầu - Xuất đầu
                ton_dau = past_pivot.get(pid, {}).get('Nhập', 0) - past_pivot.get(pid, {}).get('Xuất', 0)
                # Nhập/Xuất trong kỳ
                nhap = period_pivot.get(pid, {}).get('Nhập', 0)
                xuat = period_pivot.get(pid, {}).get('Xuất', 0)
                
                report_data.append([p.code, p.name, p.unit, ton_dau, nhap, xuat, ton_dau + nhap - xuat])
            
            df_report = pd.DataFrame(report_data, columns=["Mã HH", "Tên", "Đvt", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối"])
            
            st.dataframe(df_report, use_container_width=True, hide_index=True)
            
            st.download_button(
                label="📥 Xuất báo cáo ra Excel (.xlsx)",
                data=export_to_excel(df_report),
                file_name="BaoCaoTonKho.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )