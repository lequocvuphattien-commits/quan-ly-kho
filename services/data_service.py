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
            cleaned_data = [row[:5] for row in data[1:]]
            cleaned_data = [row + [""] * (5 - len(row)) for row in cleaned_data]

            df = pd.DataFrame(cleaned_data, columns=["date", "product_id", "type", "qty", "note"])
            
            df['product_id'] = df['product_id'].astype(str).str.strip()
            df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
            # Giữ nguyên kiểu chữ trong sheet để so sánh
            df['type'] = df['type'].astype(str).str.strip() 
            return df.values.tolist()
        return []

    def get_products(self):
        """Lấy danh mục sản phẩm"""
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            cleaned_data = [row[:5] for row in data[1:]]
            cleaned_data = [row + [""] * (5 - len(row)) for row in cleaned_data]

            df = pd.DataFrame(cleaned_data, columns=["id", "code", "name", "unit", "stock"])
            df['code'] = df['code'].astype(str).str.strip()
            return df.values.tolist()
        return []

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Tính toán tồn kho chuẩn xác với từ khóa Nhập/Xuất"""
        raw_data = self.get_history()
        if not raw_data: return 0.0, 0.0, 0.0
        
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        df['date'] = pd.to_datetime(df['date'])
        
        target_id = str(product_id).strip()
        df_prod = df[df['product_id'] == target_id]
        
        if df_prod.empty: return 0.0, 0.0, 0.0
        
        start, end = pd.to_datetime(start_date), pd.to_datetime(end_date)
        
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'Nhập']['qty'].sum() - 
                   past_data[past_data['type'] == 'Xuất']['qty'].sum())
        
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'Nhập']['qty'].sum()
        xuat = period_data[period_data['type'] == 'Xuất']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)

    def add_transaction(self, product_id, qty, trans_type, note):
        """Thêm giao dịch mới (Lưu trực tiếp chữ Nhập/Xuất)"""
        date_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        # Lưu thẳng trans_type (ví dụ: "Nhập") vào sheet
        self.sheet_transactions.append_row([date_str, str(product_id), trans_type, float(qty), str(note)])

    def add_product(self, code, name, unit):
        existing_products = self.get_products()
        next_id = len(existing_products) + 1
        self.sheet_products.append_row([str(next_id), code, name, unit, 0])

    def check_product_exists(self, code):
        products = self.get_products()
        return any(str(p[1]).strip() == str(code).strip() for p in products)

    def update_stock(self, product_code, qty, trans_type):
        """Cập nhật tồn kho dựa trên chữ Nhập/Xuất"""
        records = self.sheet_products.get_all_values()
        
        for i, row in enumerate(records):
            if i == 0: continue 
            
            if len(row) > 1 and str(row[1]).strip() == str(product_code).strip():
                current_stock = 0.0
                if len(row) > 4 and str(row[4]).strip() != "":
                    try: current_stock = float(row[4])
                    except: current_stock = 0.0
                
                try: qty_val = float(qty)
                except: qty_val = 0.0
                    
                # So sánh trực tiếp với "Nhập"
                if trans_type == "Nhập":
                    new_stock = current_stock + qty_val
                elif trans_type == "Xuất":
                    new_stock = current_stock - qty_val
                else:
                    new_stock = current_stock
                    
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                return True
        return False