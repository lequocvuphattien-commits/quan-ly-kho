import sys
import os
import uuid
import uuid
import pandas as pd

# Thêm thư mục gốc của dự án vào hệ thống để Python tìm thấy 'controllers'
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        """Lấy lịch sử giao dịch với 6 cột mới: date, product_id, product_name, type, qty, note"""
        data = self.sheet_transactions.get_all_values()
        if len(data) > 1:
            # Tăng số cột xử lý lên 6
            cleaned_data = [row[:6] for row in data[1:]]
            cleaned_data = [row + [""] * (6 - len(row)) for row in cleaned_data]

            df = pd.DataFrame(cleaned_data, columns=["date", "product_id", "product_name", "type", "qty", "note"])
            
            # Chuẩn hóa dữ liệu
            df['product_id'] = df['product_id'].astype(str).str.strip()
            df['product_name'] = df['product_name'].astype(str).str.strip()
            df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
            df['type'] = df['type'].astype(str).str.strip().str.upper()
            return df.values.tolist()
        return []

    def add_transaction(self, product_id, product_name, qty, trans_type, note):
        """Ghi đầy đủ 6 thông tin vào sheet Transactions"""
        date_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        # Đảm bảo thứ tự: Date, ID, Tên, Loại, Số lượng, Ghi chú
        self.sheet_transactions.append_row([date_str, str(product_id), str(product_name), trans_type.upper(), float(qty), note])

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Tính toán dựa trên cấu trúc 6 cột"""
        raw_data = self.get_history()
        if not raw_data: return 0.0, 0.0, 0.0
        
        # Cập nhật columns khớp với hàm get_history mới
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "product_name", "type", "qty", "note"])
        
        df['date'] = pd.to_datetime(df['date'])
        df['product_id'] = df['product_id'].astype(str).str.strip()
        df['type'] = df['type'].astype(str).str.strip()
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        
        target_id = str(product_id).strip()
        df_prod = df[df['product_id'] == target_id]
        
        if df_prod.empty: return 0.0, 0.0, 0.0
        
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'NHẬP']['qty'].sum() - 
                   past_data[past_data['type'] == 'XUẤT']['qty'].sum())
        
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'NHẬP']['qty'].sum()
        xuat = period_data[period_data['type'] == 'XUẤT']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)
    
    def get_products(self):
        """Lấy danh mục sản phẩm"""
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            # Ví dụ xử lý: cắt dữ liệu từ dòng 2 (bỏ header)
            cleaned_data = [row[:5] for row in data[1:]]
            # ... logic xử lý dữ liệu ...
            return cleaned_data
        return []
    
    def check_product_exists(self, product_code):
        """Kiểm tra mã hàng đã tồn tại trong sheet Products hay chưa"""
        products = self.get_products()
        # Dùng .strip().lower() để so sánh không phân biệt hoa thường/khoảng trắng
        # Giả định cột thứ 2 (index 1) là cột chứa Mã hàng hóa
        return any(str(p[1]).strip().lower() == str(product_code).strip().lower() for p in products)
    
    def add_product(self, code, name, unit):
        """Thêm hàng hóa mới vào Google Sheet"""
        # Tạo một ID ngẫu nhiên, cắt lấy 8 ký tự đầu và in hoa
        new_id = str(uuid.uuid4())[:8].upper()
        
        # Thêm một dòng mới vào sheet 'Products'
        # Cấu trúc: [ID, Mã HH, Tên HH, Đơn vị tính, Tồn kho ban đầu]
        self.sheet_products.append_row([
            new_id,       
            code,         
            name,         
            unit,         
            0.0           # Tồn kho ban đầu mặc định là 0
        ])
        return True