import sys
import os
import uuid
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")
        self.sheet_config = self.provider.get_sheet("Config")
        self.sheet_employees = self.provider.get_sheet("NhanVien")

    def get_history(self):
        data = self.sheet_transactions.get_all_values()
        if len(data) > 1:
            cleaned_data = [row[:7] for row in data[1:]]
            cleaned_data = [row + [""] * (7 - len(row)) for row in cleaned_data]
            df = pd.DataFrame(cleaned_data, columns=["date", "product_id", "product_name", "type", "qty", "note", "emp_name"])
            return df.values.tolist()
        return []

    def add_transaction(self, product_id, product_name, qty, trans_type, note, emp_name=""):
        date_str = pd.Timestamp.now(tz='Asia/Ho_Chi_Minh').strftime("%Y-%m-%d %H:%M:%S")
        self.sheet_transactions.append_row([date_str, str(product_id), str(product_name), trans_type.upper(), float(qty), str(note), str(emp_name)])

    def get_products(self):
        data = self.sheet_products.get_all_values()
        return data[1:] if len(data) > 1 else []

    def check_product_exists(self, product_code):
        products = self.get_products()
        return any(str(p[1]).strip().lower() == str(product_code).strip().lower() for p in products)

    def add_product(self, code, name, unit):
        new_id = str(uuid.uuid4())[:8].upper()
        self.sheet_products.append_row([new_id, code, name, unit, 0.0])
        return True

    def delete_product(self, product_id):
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id:
                self.sheet_products.delete_rows(i + 1)
                return True
        return False

    def update_product(self, product_id, new_name, new_unit):
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id:
                self.sheet_products.update_cell(i + 1, 3, new_name)
                self.sheet_products.update_cell(i + 1, 4, new_unit)
                return True
        return False

    def update_stock(self, product_code, qty, trans_type):
        records = self.sheet_products.get_all_values()
        for i, row in enumerate(records):
            if i > 0 and len(row) > 1 and str(row[1]).strip().lower() == str(product_code).strip().lower():
                current_stock = float(row[4]) if len(row) > 4 and row[4] else 0.0
                new_stock = current_stock + float(qty) if trans_type.strip().capitalize() == "Nhập" else current_stock - float(qty)
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                return True
        return False

    def get_config_options(self):
        data = self.sheet_config.get_all_values()
        return ([str(r[0]) for r in data[1:] if r[0]], [str(r[1]) for r in data[1:] if len(r)>1 and r[1]])

    def get_employees(self):
        data = self.sheet_employees.get_all_values()
        return [row[:5] + [""] * (5 - len(row)) for row in data[1:]] if len(data) > 1 else []

    def check_login(self, username, password):
        employees = self.get_employees()
        for emp in employees:
            if str(emp[0]).strip().upper() == username.strip().upper() and str(emp[4]).strip() == password:
                return {"status": True, "name": emp[1]}
        return {"status": False, "name": None}
