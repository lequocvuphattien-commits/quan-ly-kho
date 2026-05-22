import streamlit as st
import pandas as pd
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        """Lấy lịch sử giao dịch và chuẩn hóa dữ liệu"""
        data = self.sheet_transactions.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            # Chuẩn hóa cột quan trọng
            df['Mã HH'] = df['Mã HH'].astype(str).str.strip()
            df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
            df['type'] = df['type'].astype(str).str.strip().str.upper()
            return df.values.tolist()
        return []

    def get_products(self):
        """Lấy danh mục sản phẩm"""
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=data[0])
            df['Mã'] = df['Mã'].astype(str).str.strip()
            return df.values.tolist()
        return []

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Tính toán tồn kho chuẩn xác"""
        raw_data = self.get_history()
        if not raw_data: 
            return 0.0, 0.0, 0.0
        
        # Tạo DataFrame từ dữ liệu đã chuẩn hóa
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        # Chuyển đổi định dạng
        df['date'] = pd.to_datetime(df['date'])
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        df['product_id'] = df['product_id'].astype(str).str.strip()
        df['type'] = df['type'].str.upper() # Đảm bảo khớp IMPORT/EXPORT
        
        # Lọc theo mã hàng
        target_id = str(product_id).strip()
        df_prod = df[df['product_id'] == target_id]
        
        if df_prod.empty:
            return 0.0, 0.0, 0.0
        
        # Tính toán
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        # Tồn đầu
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'IMPORT']['qty'].sum() - 
                   past_data[past_data['type'] == 'EXPORT']['qty'].sum())
        
        # Nhập / Xuất trong kỳ
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'IMPORT']['qty'].sum()
        xuat = period_data[period_data['type'] == 'EXPORT']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)

    def add_transaction(self, product_id, qty, trans_type, note):
        """Thêm giao dịch mới"""
        date_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sheet_transactions.append_row([date_str, str(product_id), trans_type.upper(), qty, note])

    def add_product(self, code, name, unit):
        """Thêm sản phẩm mới"""
        self.sheet_products.append_row([code, code, name, unit, 0])

    def check_product_exists(self, code):
        """Kiểm tra mã hàng đã tồn tại chưa"""
        products = self.get_products()
        return any(str(p[1]).strip() == str(code).strip() for p in products)