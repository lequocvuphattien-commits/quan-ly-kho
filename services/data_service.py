import streamlit as st
import pandas as pd
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        """Lấy toàn bộ lịch sử giao dịch và ép kiểu Mã hàng về String"""
        data = self.sheet_transactions.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            # Ép kiểu cột Mã HH sang chuỗi và làm sạch
            df['Mã HH'] = df['Mã HH'].astype(str).str.strip()
            return df.values.tolist()
        return []

    def get_products(self):
        """Lấy danh sách sản phẩm và ép kiểu Mã hàng về String"""
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            # Ép kiểu cột Mã sang chuỗi và làm sạch
            df['Mã'] = df['Mã'].astype(str).str.strip()
            return df.values.tolist()
        return []

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Tính toán tồn kho với dữ liệu đã được làm sạch"""
        raw_data = self.get_history()
        if not raw_data: return 0.0, 0.0, 0.0
        
        # Chuyển dữ liệu lịch sử thành DataFrame
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        # Làm sạch dữ liệu trong bảng giao dịch
        df['product_id'] = df['product_id'].astype(str).str.strip()
        df['date'] = pd.to_datetime(df['date'])
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        
        # Ép kiểu biến đầu vào về dạng String để so sánh
        target_id = str(product_id).strip()
        
        # Lọc dữ liệu khớp với mã hàng
        df_prod = df[df['product_id'] == target_id]
        
        # Tính toán
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'IMPORT']['qty'].sum() - 
                   past_data[past_data['type'] == 'EXPORT']['qty'].sum())
        
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'IMPORT']['qty'].sum()
        xuat = period_data[period_data['type'] == 'EXPORT']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)

    def add_transaction(self, product_id, qty, trans_type, note):
        date_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        # Ghi vào sheet
        self.sheet_transactions.append_row([date_str, str(product_id), trans_type, qty, note])