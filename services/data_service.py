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
        if not raw_data: return 0.0, 0.0, 0.0
        
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        # Làm sạch dữ liệu
        df['date'] = pd.to_datetime(df['date'])
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        df['product_id'] = df['product_id'].astype(str).str.strip()
        
        # Lọc (Bây giờ nó sẽ khớp vì product_id truyền vào đã là "H01")
        df_prod = df[df['product_id'] == str(product_id).strip()]
        
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'IMPORT']['qty'].sum() - 
                   past_data[past_data['type'] == 'EXPORT']['qty'].sum())
        
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'IMPORT']['qty'].sum()
        xuat = period_data[period_data['type'] == 'EXPORT']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)