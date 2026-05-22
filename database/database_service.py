import sqlite3
import os

class DatabaseService:
    """
    Tiện ích quản lý kết nối SQLite chung cho ứng dụng.
    Chỉ dùng để lấy connection, không chứa logic nghiệp vụ.
    """
    def __init__(self, db_name="inventory.db"):
        # Lấy đường dẫn thư mục database
        self.db_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "database", db_name)

    def get_connection(self):
        """Trả về đối tượng kết nối database"""
        return sqlite3.connect(self.db_path)
    
    def get_product_stats_by_date(self, product_id, start_date, end_date):
        return self.provider.get_product_stats_by_date(product_id, start_date, end_date)

# Tạo một instance duy nhất để các nơi khác có thể import
db_service = DatabaseService()