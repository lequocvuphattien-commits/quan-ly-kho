import sys
import os
import uuid
import uuid
import pandas as pd
import datetime

# Thêm thư mục gốc của dự án vào hệ thống để Python tìm thấy 'controllers'
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")
        self.sheet_config = self.provider.get_sheet("Config")

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
        # Ép múi giờ về Việt Nam (UTC+7)
        date_str = pd.Timestamp.now(tz='Asia/Ho_Chi_Minh').strftime("%Y-%m-%d %H:%M:%S")
        # Đảm bảo thứ tự: Date, ID, Tên, Loại, Số lượng, Ghi chú
        self.sheet_transactions.append_row([date_str, str(product_id), str(product_name), trans_type.upper(), float(qty), str(note)])

    def get_config_options(self):
        """Đọc danh sách Kho Nhập (Cột A) và Kho Xuất (Cột B) từ sheet Config"""
        try:
            data = self.sheet_config.get_all_values()
            if len(data) <= 1:
                return [], [] # Trả về rỗng nếu sheet chưa có dữ liệu (chỉ có header)
            
            kho_nhap = []
            kho_xuat = []
            
            # Duyệt từ dòng thứ 2 trở đi (bỏ qua dòng tiêu đề)
            for row in data[1:]:
                # Cột A (index 0) là Kho Nhập
                if len(row) > 0 and str(row[0]).strip() != "":
                    kho_nhap.append(str(row[0]).strip())
                # Cột B (index 1) là Kho Xuất
                if len(row) > 1 and str(row[1]).strip() != "":
                    kho_xuat.append(str(row[1]).strip())
                    
            return kho_nhap, kho_xuat
        except Exception as e:
            print(f"Lỗi đọc config: {e}")
            return [], []

    def add_config_option(self, option_type, new_value):
        """Thêm một kho mới vào sheet Config"""
        try:
            # Nếu loại là 'Nhập' thì ghi vào Cột 1 (A), 'Xuất' thì ghi Cột 2 (B)
            col_index = 1 if option_type == "Nhập" else 2
            
            # Lấy toàn bộ dữ liệu của cột đó để đếm xem có bao nhiêu dòng rồi
            col_values = self.sheet_config.col_values(col_index)
            
            # Tính ra vị trí dòng trống tiếp theo
            next_row = len(col_values) + 1
            
            # Ghi giá trị mới vào ô trống đó
            self.sheet_config.update_cell(next_row, col_index, new_value)
            return True
        except Exception as e:
            print(f"Lỗi thêm config: {e}")
            return False

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
    
    # Cập nhật và xóa sản phẩm dựa trên Mã hàng (cột Mã HH - index 1)
    def update_product(self, product_id, new_name, new_unit):
        """Cập nhật thông tin sản phẩm dựa trên ID (cột A)"""
        # Tìm dòng có product_id trùng với cột Mã (giả sử cột Mã là cột B - index 1)
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id: # Index 1 là cột Mã
                self.sheet_products.update_cell(i + 1, 3, new_name) # Cột Tên (C)
                self.sheet_products.update_cell(i + 1, 4, new_unit) # Cột Đvt (D)
                return True
        return False
    # Xóa sản phẩm dựa trên Mã hàng (cột Mã HH - index 1)
    def delete_product(self, product_id):
        """Xóa hàng hóa dựa trên ID"""
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id:
                self.sheet_products.delete_rows(i + 1)
                return True
        return False
    
    def update_stock(self, product_code, qty, trans_type):
        """Cập nhật số lượng tồn kho trực tiếp vào sheet Products"""
        # Lấy toàn bộ dữ liệu từ sheet Products
        records = self.sheet_products.get_all_values()
        
        for i, row in enumerate(records):
            # i > 0 để bỏ qua dòng tiêu đề. Giả định cột Mã HH là cột thứ 2 (index 1)
            if i > 0 and len(row) > 1 and str(row[1]).strip().lower() == str(product_code).strip().lower():
                # Lấy số lượng tồn hiện tại ở cột thứ 5 (index 4)
                current_stock = float(row[4]) if len(row) > 4 and row[4] else 0.0
                
                # Tính toán lại tồn kho
                if trans_type.strip().capitalize() == "Nhập":
                    new_stock = current_stock + float(qty)
                else: # Xuất
                    new_stock = current_stock - float(qty)
                
                # Cập nhật ô trong Google Sheets (Dòng i+1 do gspread đếm từ 1, Cột 5)
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                return True
                
        return False # Trả về False nếu không tìm thấy mã hàng
    