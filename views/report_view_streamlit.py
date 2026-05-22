import streamlit as st
import pandas as pd
from controllers.transaction_controller import TransactionController
from controllers.product_controller import ProductController

def show_report():
    st.subheader("Báo cáo tồn kho")
    
    p_controller = ProductController()
    t_controller = TransactionController()
    
    # Bộ lọc ngày tháng
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Từ ngày").strftime("%Y-%m-%d 00:00:00")
    end_date = col2.date_input("Đến ngày").strftime("%Y-%m-%d 23:59:59")
    
    if st.button("Lọc báo cáo"):
        products = p_controller.get_all_products()
        report_data = []
        
        for p in products:
            dau, nhap, xuat = t_controller.get_product_stats_by_date(p.id, start_date, end_date)
            cuoi = dau + nhap - xuat
            report_data.append({
                "Mã HH": p.code,
                "Tên": p.name,
                "Đvt": p.unit,
                "Tồn Đầu": dau,
                "Nhập": nhap,
                "Xuất": xuat,
                "Tồn Cuối": cuoi
            })
        
        df = pd.DataFrame(report_data)
        st.dataframe(df, use_container_width=True)
        
        # Nút xuất Excel
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Tải xuống CSV", data=csv, file_name="BaoCaoTonKho.csv")