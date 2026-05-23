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