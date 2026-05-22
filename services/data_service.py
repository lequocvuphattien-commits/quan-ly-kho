import streamlit as st
import pandas as pd
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        """Lấy toàn bộ lịch sử giao dịch"""
        data = self.sheet_transactions.get_all_values()
        return data[1:] if len(data) > 1 else []

    def get_products(self):
        """Lấy danh sách hàng hóa"""
        data = self.sheet_products.get_all_values()
        return data[1:] if len(data) > 1 else []

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        raw_data = self.get_history()
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        # 1. Ép kiểu và làm sạch
        df['product_id'] = df['product_id'].astype(str).str.strip()
        target_id = str(product_id).strip()
        
        # 2. DEBUG: Xem danh sách các mã hàng có trong hệ thống
        unique_ids = df['product_id'].unique()
        st.write(f"Các mã hàng hiện có trong Sheets: {unique_ids}")
        st.write(f"Mã hàng bạn đang chọn: '{target_id}'")
        
        # 3. Lọc
        df_prod = df[df['product_id'] == target_id]
        
        if df_prod.empty:
            st.error(f"Lỗi: Không tìm thấy mã '{target_id}' trong danh sách. Hãy kiểm tra xem tên cột có bị sai hoặc mã hàng có thừa khoảng trắng không!")
            return 0.0, 0.0, 0.0
            
        # Tính toán (nếu tìm thấy dữ liệu)
        df['date'] = pd.to_datetime(df['date'])
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        
        # ... (phần tính toán tồn kho giữ nguyên)
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'IMPORT']['qty'].sum() - 
                   past_data[past_data['type'] == 'EXPORT']['qty'].sum())
        
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'IMPORT']['qty'].sum()
        xuat = period_data[period_data['type'] == 'EXPORT']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)