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
            return True
        except Exception as e:
            print(f"Lỗi khi xử lý giao dịch: {e}")
            return False

    def get_transaction_history(self):
        """Lấy danh sách toàn bộ lịch sử giao dịch"""
        return self.service.get_history()

    def undo_transaction(self, trans_id):
        """Hủy giao dịch (cần được cài đặt trong DataService)"""
        try:
            return self.service.undo_transaction(trans_id)
        except Exception as e:
            print(f"Lỗi khi hủy giao dịch: {e}")
            return False

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """
        Tính toán tồn đầu kỳ, tổng nhập, tổng xuất trong khoảng thời gian.
        Đây là hàm quan trọng cho Báo cáo tồn kho.
        """
        try:
            return self.service.get_product_stats_by_date(product_id, start_date, end_date)
        except AttributeError:
            print("Lỗi: Hàm get_product_stats_by_date chưa được định nghĩa trong DataService!")
            return 0.0, 0.0, 0.0
        except Exception as e:
            print(f"Lỗi khi tính toán báo cáo: {e}")
            return 0.0, 0.0, 0.0