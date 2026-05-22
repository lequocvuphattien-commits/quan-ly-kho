import streamlit as st
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
import datetime

class DataService:
    def __init__(self, mode="ONLINE"):
        self.mode = mode
        if mode == "ONLINE":
            try:
                secrets = st.secrets["gcp"]["json_key"]
                key_dict = json.loads(secrets)
                scope = [
                    "https://spreadsheets.google.com/feeds", 
                    'https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"
                ]
                creds = ServiceAccountCredentials.from_json_keyfile_dict(key_dict, scope)
                self.client = gspread.authorize(creds)
                
                # SỬA LỖI 1: Dùng ID thay vì tên file để kết nối chính xác 100%
                FILE_ID = "1Eu0opyFfWiDOWGAq4Ox2cg-UF-_nxh8XVvvoNvF7DLo"
                self.spreadsheet = self.client.open_by_key(FILE_ID)
                
                self.sheet_products = self.spreadsheet.worksheet("Products")
                self.sheet_transactions = self.spreadsheet.worksheet("Transactions")
            except Exception as e:
                st.error(f"Lỗi kết nối Google Sheets: {e}")
                raise e

    def get_products(self):
        data = self.sheet_products.get_all_values()
        # Loại bỏ hàng tiêu đề và các hàng trống
        return [row for row in data[1:] if any(row)]

    def add_product(self, code, name, unit):
        current_data = self.get_products()
        new_id = len(current_data) + 1
        self.sheet_products.append_row([str(new_id), code.upper(), name, unit, 0])

    def check_product_exists(self, code):
        products = self.get_products()
        for p in products:
            if len(p) > 1 and p[1].upper() == code.upper():
                return True
        return False

    def add_transaction(self, product_code, quantity, trans_type, note):
        date_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sheet_transactions.append_row([date_now, product_code, trans_type, float(quantity), note])

    # SỬA LỖI 2: Thêm hàm cập nhật tồn kho vào Google Sheets
    def update_stock(self, product_code, quantity, trans_type):
        all_values = self.sheet_products.get_all_values()
        for i in range(1, len(all_values)):
            # Cột 1 là Mã HH
            if str(all_values[i][1]).upper() == str(product_code).upper():
                current_stock = float(all_values[i][4]) if (all_values[i][4] and all_values[i][4].replace('.','',1).isdigit()) else 0
                
                if trans_type == "IMPORT":
                    new_stock = current_stock + float(quantity)
                else:
                    new_stock = current_stock - float(quantity)
                
                # Cập nhật ô vào cột thứ 5 (Tồn)
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                return True
        return False

    def get_history(self):
        data = self.sheet_transactions.get_all_values()
        return [row for row in data[1:] if any(row)]