import streamlit as st
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
import datetime

class DataService:
    def __init__(self, mode="ONLINE"):
        self.mode = mode
        self.sheet_products = None
        self.sheet_transactions = None
        
        if mode == "ONLINE":
            try:
                secrets = st.secrets["gcp"]["json_key"]
                key_dict = json.loads(secrets)
                scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_dict(key_dict, scope)
                self.client = gspread.authorize(creds)
                FILE_ID = "1Eu0opyFfWiDOWGAq4Ox2cg-UF-_nxh8XVvvoNvF7DLo"
                self.spreadsheet = self.client.open_by_key(FILE_ID)
                self.sheet_products = self.spreadsheet.worksheet("Products")
                self.sheet_transactions = self.spreadsheet.worksheet("Transactions")
            except Exception as e:
                st.error(f"Lỗi kết nối: {e}")

    # Bỏ cache ở đây vì ta muốn lấy dữ liệu mới nhất từ Sheet
    def get_products(self):
        if self.sheet_products is None: return []
        data = self.sheet_products.get_all_values()
        return data[1:] if len(data) > 1 else []

    def add_product(self, code, name, unit):
        if self.sheet_products:
            current_data = self.get_products()
            new_id = len(current_data) + 1
            self.sheet_products.append_row([str(new_id), code.upper(), name, unit, 0])

    def check_product_exists(self, code):
        products = self.get_products()
        for p in products:
            if len(p) > 1 and p[1].upper() == code.upper(): return True
        return False

    def add_transaction(self, product_code, quantity, trans_type, note):
        if self.sheet_transactions:
            date_now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.sheet_transactions.append_row([date_now, product_code, trans_type, float(quantity), note])

    def update_stock(self, product_code, quantity, trans_type):
        """Cập nhật tồn kho có in log ra màn hình để kiểm tra"""
        all_values = self.sheet_products.get_all_values()
        st.write(f"Đang tìm mã hàng: '{product_code}' trong {len(all_values)} dòng...")
        
        found = False
        for i in range(1, len(all_values)):
            row = all_values[i]
            # In ra mã hàng đang duyệt để bạn đối chiếu
            # st.write(f"Đang so sánh '{row[1]}' với '{product_code}'") 
            
            if str(row[1]).strip().upper() == str(product_code).strip().upper():
                found = True
                current_stock = float(row[4]) if (row[4] and str(row[4]).replace('.','',1).isdigit()) else 0
                
                if trans_type == "IMPORT":
                    new_stock = current_stock + float(quantity)
                else:
                    new_stock = current_stock - float(quantity)
                
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                st.success(f"Đã cập nhật dòng {i+1}, tồn mới: {new_stock}")
                return True
        
        if not found:
            st.error(f"LỖI: Không tìm thấy mã hàng '{product_code}' trong Sheet!")
        return False

    def get_history(self):
        if self.sheet_transactions:
            data = self.sheet_transactions.get_all_values()
            return data[1:] if len(data) > 1 else []
        return []
    
    def get_product_stats_by_date(self, product_id, start_date, end_date):
        # Lấy dữ liệu thô từ hàm get_history()
        raw_data = self.get_history()
        if not raw_data: return 0.0, 0.0, 0.0
        
        # ĐẶT TÊN CỘT CHÍNH XÁC (Dựa theo thứ tự trong Google Sheet của bạn)
        # Giả sử: Cột 0: Ngày, Cột 1: Mã hàng, Cột 2: Loại, Cột 3: Số lượng
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        # Chuyển đổi kiểu dữ liệu
        df['date'] = pd.to_datetime(df['date'])
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        df['product_id'] = df['product_id'].astype(str).str.strip()
        
        # Lọc theo mã hàng (so sánh chuỗi đã xóa khoảng trắng)
        df_prod = df[df['product_id'] == str(product_id).strip()]
        
        # Lọc thời gian
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        # Tính toán
        past = df_prod[df_prod['date'] < start]
        ton_dau = (past[past['type'] == 'IMPORT']['qty'].sum() - 
                   past[past['type'] == 'EXPORT']['qty'].sum())
        
        period = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period[period['type'] == 'IMPORT']['qty'].sum()
        xuat = period[period['type'] == 'EXPORT']['qty'].sum()
        
        return float(ton_dau), float(nhap), float(xuat)
        