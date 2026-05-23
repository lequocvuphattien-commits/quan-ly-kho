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
            # Cắt lấy đúng 5 cột đầu tiên của mỗi dòng, đề phòng sheet có cột rác
            cleaned_data = [row[:5] for row in data[1:]]
            # Nếu dòng nào bị thiếu cột (do xóa nhầm), tự động bù khoảng trắng cho đủ 5
            cleaned_data = [row + [""] * (5 - len(row)) for row in cleaned_data]

            df = pd.DataFrame(cleaned_data, columns=["date", "product_id", "type", "qty", "note"])
            
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
            # Cắt lấy đúng 5 cột đầu tiên của mỗi dòng, đề phòng sheet có cột rác
            cleaned_data = [row[:5] for row in data[1:]]
            cleaned_data = [row + [""] * (5 - len(row)) for row in cleaned_data]

            df = pd.DataFrame(cleaned_data, columns=["id", "code", "name", "unit", "stock"])
            df['code'] = df['code'].astype(str).str.strip()
            return df.values.tolist()
        return []

    def get_product_map(self):
        """Dùng cho Selectbox: lấy thông tin nhanh để hiển thị Tồn"""
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=["id", "code", "name", "unit", "stock"])
            df['stock'] = pd.to_numeric(df['stock'], errors='coerce').fillna(0)
            return {str(row['code']).strip(): {'name': row['name'], 'unit': row['unit'], 'stock': row['stock']} for _, row in df.iterrows()}
        return {}

    def add_transaction(self, product_id, qty, trans_type, note):
        date_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sheet_transactions.append_row([date_str, str(product_id), trans_type, float(qty), str(note)])

    def update_stock(self, product_code, qty, trans_type):
        records = self.sheet_products.get_all_values()
        for i, row in enumerate(records):
            if i > 0 and len(row) > 1 and str(row[1]).strip() == str(product_code).strip():
                current = float(row[4]) if len(row) > 4 and row[4] else 0.0
                new_stock = current + float(qty) if trans_type == "Nhập" else current - float(qty)
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                return True
        return False
    
    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Tính tồn đầu kỳ và nhập xuất trong kỳ"""
        raw_data = self.get_history()
        if not raw_data: return 0.0, 0.0, 0.0
        
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        # Chuẩn hóa dữ liệu cực kỳ quan trọng
        df['date'] = pd.to_datetime(df['date'])
        df['product_id'] = df['product_id'].astype(str).str.strip()
        df['type'] = df['type'].astype(str).str.strip()
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        
        target_id = str(product_id).strip()
        df_prod = df[df['product_id'] == target_id]
        
        if df_prod.empty: return 0.0, 0.0, 0.0
        
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        # Tồn đầu kỳ = (Tổng nhập trước start) - (Tổng xuất trước start)
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'Nhập']['qty'].sum() - 
                   past_data[past_data['type'] == 'Xuất']['qty'].sum())
        
        # Trong kỳ
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'Nhập']['qty'].sum()
        xuat = period_data[period_data['type'] == 'Xuất']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)
    
    def check_product_exists(self, product_code):
        """Kiểm tra mã hàng đã tồn tại trong sheet Products hay chưa"""
        products = self.get_products()
        # Dùng .strip().lower() để so sánh không phân biệt hoa thường/khoảng trắng
        return any(str(p[1]).strip().lower() == str(product_code).strip().lower() for p in products)
    
    # Mở file services/data_service.py và thêm hàm này vào class DataService
    def add_product(self, code, name, unit):
        """Thêm hàng hóa mới vào Google Sheet"""
        # Giả sử bạn dùng thư viện gspread để tương tác với Google Sheets
        # Dòng này thêm một hàng mới vào sheet 'Products'
        self.sheet_products.append_row([
            "",          # ID (để trống nếu sheet tự tăng)
            code,        # Mã
            name,        # Tên
            unit,        # Đơn vị tính
            0.0          # Tồn kho ban đầu
        ])
        return True