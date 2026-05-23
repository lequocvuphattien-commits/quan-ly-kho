import streamlit as st
import pandas as pd
from controllers.transaction_controller import TransactionController
from controllers.product_controller import ProductController

def show_report():
    st.subheader("Báo cáo tồn kho")
    
    p_controller = ProductController()
    t_controller = TransactionController()
    
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Từ ngày").strftime("%Y-%m-%d 00:00:00")
    end_date = col2.date_input("Đến ngày").strftime("%Y-%m-%d 23:59:59")
    
    if st.button("Lọc báo cáo", type="primary"):
        with st.spinner('Đang tính toán dữ liệu tốc độ cao...'):
            products = p_controller.get_all_products()
            
            if not products:
                st.warning("Không tìm thấy hàng hóa trong danh mục!")
                return
                
            all_history_df = t_controller.get_transaction_history()
            
            if not all_history_df:
                st.info("Chưa có giao dịch nào trong lịch sử.")
                return

            df_hist = pd.DataFrame(all_history_df, columns=["date", "product_id", "type", "qty", "note"])
            df_hist['date'] = pd.to_datetime(df_hist['date'])
            df_hist['qty'] = pd.to_numeric(df_hist['qty'], errors='coerce').fillna(0)
            # TỐI ƯU: Loại bỏ khoảng trắng thừa ở cột 'type' để so sánh chính xác
            df_hist['type'] = df_hist['type'].astype(str).str.strip()
            
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)
            
            report_data = []
            
            for p in products:
                # Ép kiểu Mã hàng về chuỗi và strip để so sánh chính xác
                df_prod = df_hist[df_hist['product_id'] == str(p.code).strip()]
                
                # Tính Tồn đầu (Type phải là "Nhập" hoặc "Xuất" chính xác)
                past = df_prod[df_prod['date'] < start]
                ton_dau = past[past['type'] == 'Nhập']['qty'].sum() - past[past['type'] == 'Xuất']['qty'].sum()
                
                # Tính Nhập/Xuất trong kỳ
                period = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
                nhap = period[period['type'] == 'Nhập']['qty'].sum()
                xuat = period[period['type'] == 'Xuất']['qty'].sum()
                cuoi = ton_dau + nhap - xuat
                
                report_data.append({
                    "Mã HH": p.code,
                    "Tên": p.name,
                    "Đvt": p.unit,
                    "Tồn Đầu": float(ton_dau),
                    "Nhập": float(nhap),
                    "Xuất": float(xuat),
                    "Tồn Cuối": float(cuoi)
                })
            
            df_report = pd.DataFrame(report_data)
            st.dataframe(df_report, use_container_width=True, hide_index=True)
            
            csv = df_report.to_csv(index=False).encode('utf-8')
            st.download_button(label="Tải xuống báo cáo (CSV)", data=csv, file_name="BaoCaoTonKho.csv", mime="text/csv")