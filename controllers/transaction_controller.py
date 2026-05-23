import streamlit as st
import sys
import os

# Đảm bảo import đúng DataService
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from services.data_service import DataService

class TransactionController:
    def __init__(self):
        self.service = DataService(mode="ONLINE")

    def process_transaction(self, product_id, trans_type, quantity, note):
        """
        Xử lý giao dịch: Kiểm tra tồn kho (nếu là Xuất) và cập nhật dữ liệu.
        """
        try:
            # 1. Nếu là XUẤT, bắt buộc kiểm tra tồn kho trước
            if trans_type == "Xuất":
                products = self.service.get_products()
                # Tìm sản phẩm: giả sử cấu trúc row là [id, code, name, unit, stock]
                prod = next((p for p in products if str(p[1]).strip() == str(product_id).strip()), None)
                
                if not prod:
                    return "ERROR_NOT_FOUND"
                
                current_stock = float(prod[4])
                if float(quantity) > current_stock:
                    return "ERROR_INSUFFICIENT_STOCK"

            # 2. Ghi vào lịch sử và cập nhật tồn kho
            self.service.add_transaction(product_id, quantity, trans_type, note)
            self.service.update_stock(product_id, quantity, trans_type)
            
            # 3. Xóa cache để báo cáo cập nhật ngay lập tức
            st.cache_data.clear()
            return True
            
        except Exception as e:
            print(f"Lỗi khi xử lý giao dịch: {e}")
            return False

    def get_transaction_history(self):
        """Lấy toàn bộ lịch sử giao dịch từ Google Sheets"""
        return self._fetch_history_cached()

    @st.cache_data(ttl=60, show_spinner=False)
    def _fetch_history_cached(_self):
        """Hàm nội bộ để lấy và lưu đệm dữ liệu"""
        return _self.service.get_history()

    def undo_transaction(self, trans_id):
        """Hủy giao dịch (cần DataService hỗ trợ hàm này)"""
        try:
            result = self.service.undo_transaction(trans_id)
            if result:
                st.cache_data.clear()
            return result
        except Exception as e:
            print(f"Lỗi khi hủy giao dịch: {e}")
            return False

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Lấy thống kê báo cáo theo khoảng thời gian"""
        try:
            return self.service.get_product_stats_by_date(product_id, start_date, end_date)
        except Exception as e:
            print(f"Lỗi tính toán báo cáo: {e}")
            return 0.0, 0.0, 0.0