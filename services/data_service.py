import streamlit as st
import pandas as pd
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        """Lấy toàn bộ lịch sử giao dịch"""
        data = self.sheet_transactions.get_all_values()
        # Bỏ qua dòng tiêu đề (dòng 0), trả về dữ liệu nếu có
        return data[1:] if len(data) > 1 else []

    def get_products(self):
        """Lấy danh sách hàng hóa"""
        data = self.sheet_products.get_all_values()
        return data[1:] if len(data) > 1 else []

    import streamlit as st
import pandas as pd
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")

    def get_history(self):
        """Lấy toàn bộ lịch sử giao dịch"""
        data = self.sheet_transactions.get_all_values()
        return data[1:] if len(data) > 1 else []
    
    def get_products(self):
        """Lấy danh sách hàng hóa"""
        if self.sheet_products:
            data = self.sheet_products.get_all_values()
            return data[1:] if len(data) > 1 else []
        return []

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        raw_data = self.get_history()
        if not raw_data: 
            st.warning("Không có dữ liệu trong Transactions!")
            return 0.0, 0.0, 0.0
        
        # 1. Tạo DataFrame và DEBUG dữ liệu thô
        df = pd.DataFrame(raw_data, columns=["date", "product_id", "type", "qty", "note"])
        
        st.write("--- DEBUG: DỮ LIỆU ĐỌC TỪ GOOGLE SHEET ---")
        st.dataframe(df.head()) # Xem 5 dòng đầu
        st.write("Mã hàng bạn chọn để lọc:", product_id)

        # 2. Làm sạch dữ liệu
        df['date'] = pd.to_datetime(df['date'])
        df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0)
        df_prod = df[df['product_id'].astype(str).str.strip() == str(product_id).strip()]
        
        # 3. Lọc theo mã hàng và DEBUG kết quả lọc
        df_prod = df[df['product_id'] == str(product_id).strip()]
        st.write("--- DEBUG: DỮ LIỆU SAU KHI LỌC MÃ HÀNG ---")
        st.dataframe(df_prod)

        # 4. Tính toán
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        past_data = df_prod[df_prod['date'] < start]
        ton_dau = (past_data[past_data['type'] == 'IMPORT']['qty'].sum() - 
                   past_data[past_data['type'] == 'EXPORT']['qty'].sum())
        
        period_data = df_prod[(df_prod['date'] >= start) & (df_prod['date'] <= end)]
        nhap = period_data[period_data['type'] == 'IMPORT']['qty'].sum()
        xuat = period_data[period_data['type'] == 'EXPORT']['qty'].sum()
        
        st.write(f"Tồn đầu: {ton_dau} | Nhập: {nhap} | Xuất: {xuat}")
        
        return float(ton_dau), float(nhap), float(xuat)

    def add_transaction(self, product_id, qty, trans_type, note):
        # Logic ghi vào Google Sheets
        date_str = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sheet_transactions.append_row([date_str, product_id, trans_type, qty, note])