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
            # SỬA LỖI Ở ĐÂY: Ép tên cột tiếng Anh ngay từ đầu, bỏ qua data[0] (tiếng Việt)
            df = pd.DataFrame(data[1:], columns=["date", "product_id", "type", "qty", "note"])
            
            # Chuẩn hóa bằng tên cột tiếng Anh
            df['product_id'] = df['product_id'].astype(str).str.strip()
            df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
            df['type'] = df['type'].astype(str).str.strip().str.upper()
            return df.values.tolist()
        return []

    def get_products(self):
        """Lấy danh mục sản phẩm"""
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            # Ép tên cột tiếng Anh cho an toàn
            df = pd.DataFrame(data[1:], columns=["id", "code", "name", "unit", "stock"])
            df['code'] = df['code'].astype(str).str.strip()
            return df.values.tolist()
        return []

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Tính toán tồn kho chuẩn xác"""
        raw_data = self.get_history()
        if not raw_data: 
            return 0.0, 0.0, 0.0
        
        # Tạo DataFrame từ dữ liệu đã chuẩn hóa (Tên cột đã là tiếng Anh)
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        # Chuyển đổi định dạng ngày tháng
        df['date'] = pd.to_datetime(df['date'])
        
        # Lọc theo mã hàng (ép kiểu để so sánh an toàn)
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
        self.sheet_transactions.append_row([date_str, str(product_id), str(trans_type).upper(), float(qty), str(note)])

    def add_product(self, code, name, unit):
        """Thêm sản phẩm mới"""
        self.sheet_products.append_row(["", code, name, unit, 0]) # Giữ ID trống để công thức tính hoặc bỏ qua

    def check_product_exists(self, code):
        """Kiểm tra mã hàng đã tồn tại chưa"""
        products = self.get_products()
        # Ở hàm get_products, cột mã (code) đang nằm ở index 1
        return any(str(p[1]).strip() == str(code).strip() for p in products)