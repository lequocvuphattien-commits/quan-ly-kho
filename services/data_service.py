import streamlit as st
import pandas as pd
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        data = self.sheet_transactions.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=["date", "product_id", "type", "qty", "note"])
            df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
            return df.values.tolist()
        return []

    def get_products(self):
        data = self.sheet_products.get_all_values()
        if len(data) > 1:
            df = pd.DataFrame(data[1:], columns=["id", "code", "name", "unit", "stock"])
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