import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

class GoogleProvider:
    def __init__(self):
        self.scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        
        # 1. Lấy thông tin xác thực từ secrets (phần [gcp])
        self.credentials = Credentials.from_service_account_info(
            st.secrets["gcp"], scopes=self.scopes
        )
        
        # 2. Khởi tạo kết nối gspread
        self.gc = gspread.authorize(self.credentials)
        
        # 3. Lấy SPREADSHEET_ID từ secrets thay vì dùng tên file "DataKho"
        # Cách này giúp ứng dụng không bị lỗi nếu bạn đổi tên file trên Google Drive
        self.spreadsheet_id = st.secrets["SPREADSHEET_ID"]
        
        # 4. Mở file bằng ID (cách này chuẩn và ổn định nhất)
        self.sh = self.gc.open_by_key(self.spreadsheet_id) 

    def get_sheet(self, sheet_name):
        """Lấy worksheet theo tên tab."""
        return self.sh.worksheet(sheet_name)