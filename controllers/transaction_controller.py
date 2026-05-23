import streamlit as st
import sys
import os

# Thêm đường dẫn dự án vào sys.path để import các module dễ dàng
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from services.data_service import DataService

class TransactionController:
    def __init__(self):
        # Khởi tạo dịch vụ dữ liệu
        self.service = DataService(mode="ONLINE")

    def process_transaction(self, product_id, trans_type, quantity, note):
        """
        Xử lý giao dịch: Ghi vào lịch sử và cập nhật tồn kho ngay lập tức.
        """
        try:
            # 1. Ghi vào lịch sử giao dịch
            self.service.add_transaction(product_id, quantity, trans_type, note)
            # 2. Cập nhật tồn kho tự động
            self.service.update_stock(product_id, quantity, trans_type)
            
            # TỐI ƯU HÓA: Xóa toàn bộ bộ nhớ đệm ngay khi có dữ liệu mới
            # Điều này đảm bảo báo cáo tồn kho luôn hiển thị con số mới nhất 
            # mà không bị kẹt lại dữ liệu cũ của 60 giây trước.
            st.cache_data.clear()
            
            return True
        except Exception as e:
            print(f"Lỗi khi xử lý giao dịch: {e}")
            return False

    def get_transaction_history(self):
        """Lấy danh sách toàn bộ lịch sử giao dịch qua bộ nhớ đệm"""
        # Gọi hàm nội bộ đã được gắn Cache
        return self._fetch_history_cached()

    # TỐI ƯU HÓA TỐC ĐỘ: Bật bộ nhớ đệm (Cache) trong 60 giây
    # Dấu '_' trước chữ 'self' giúp Streamlit hiểu đây là class method và không bị lỗi hash
    @st.cache_data(ttl=60, show_spinner=False)
    def _fetch_history_cached(_self):
        """Hàm nội bộ để lấy và lưu đệm dữ liệu từ Google Sheets"""
        return _self.service.get_history()

    def undo_transaction(self, trans_id):
        """Hủy giao dịch (cần được cài đặt trong DataService)"""
        try:
            result = self.service.undo_transaction(trans_id)
            if result:
                st.cache_data.clear() # Xóa cache nếu hủy thành công
            return result
        except Exception as e:
            print(f"Lỗi khi hủy giao dịch: {e}")
            return False

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """
        Tính toán tồn đầu kỳ, tổng nhập, tổng xuất trong khoảng thời gian.
        """
        try:
            return self.service.get_product_stats_by_date(product_id, start_date, end_date)
        except AttributeError:
            print("Lỗi: Hàm get_product_stats_by_date chưa được định nghĩa trong DataService!")
            return 0.0, 0.0, 0.0
        except Exception as e:
            print(f"Lỗi khi tính toán báo cáo: {e}")
            return 0.0, 0.0, 0.0
        
    def process_transaction(self, product_id, trans_type, quantity, note):
        """
        Xử lý giao dịch: Kiểm tra tồn kho trước khi xuất.
        """
        try:
            # 1. Nếu là XUẤT, kiểm tra tồn kho trước
            if trans_type == "Xuất":
                # Lấy tồn kho hiện tại
                products = self.service.get_products()
                # Tìm sản phẩm tương ứng
                prod = next((p for p in products if str(p[1]).strip() == str(product_id).strip()), None)
                
                if prod:
                    current_stock = float(prod[4]) # Cột 4 là tồn
                    if float(quantity) > current_stock:
                        return "ERROR_INSUFFICIENT_STOCK"
                else:
                    return "ERROR_NOT_FOUND"

            # 2. Nếu đủ điều kiện, ghi vào lịch sử và cập nhật
            self.service.add_transaction(product_id, quantity, trans_type, note)
            self.service.update_stock(product_id, quantity, trans_type)
            st.cache_data.clear()
            return True
        except Exception as e:
            print(f"Lỗi khi xử lý giao dịch: {e}")
            return False