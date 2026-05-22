import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

class GoogleProvider:
    def __init__(self):
        self.scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        # Lấy thông tin từ secrets trên Streamlit Cloud
        self.credentials = Credentials.from_service_account_info(
            st.secrets["gcp"], scopes=self.scopes
        )
        self.gc = gspread.authorize(self.credentials)
        # Thay đổi tên file Google Sheet của bạn ở đây nếu cần
        self.sh = self.gc.open("quan-ly_kho") 

    def get_sheet(self, sheet_name):
        return self.sh.worksheet(sheet_name)