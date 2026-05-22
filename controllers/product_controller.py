from services.data_service import DataService
from models.product_model import Product

class ProductController:
    """
    Controller xử lý logic nghiệp vụ cho Hàng hóa (Product).
    Đóng vai trò trung gian giữa View và DataService.
    """
    def __init__(self):
        # Kết nối tới tổng đài dữ liệu (DataService)
        self.service = DataService()

    def get_all_products(self):
        """Lấy danh sách tất cả hàng hóa từ DataService và chuyển đổi sang model"""
        rows = self.service.get_products()
        products = []
        if rows:
            for row in rows:
                # row[0]: id, row[1]: code, row[2]: name, row[3]: unit, row[4]: stock
                products.append(Product(id=row[0], code=row[1], name=row[2], unit=row[3], stock=row[4]))
        return products

    def add_product(self, code, name, unit):
        # Kiểm tra trùng mã trước khi thêm
        if self.service.check_product_exists(code):
            return False, f"Mã hàng hóa '{code}' đã tồn tại!"
        
        self.service.add_product(code, name, unit)
        return True, "Thêm hàng hóa thành công!"

    def delete_product(self, product_id):
        """Xóa hàng hóa"""
        self.service.delete_product(product_id)

    def update_product(self, product_id, name, unit):
        """Cập nhật thông tin hàng hóa"""
        self.service.update_product(product_id, name, unit)

    def get_product_stats_by_date(self, product_id, start_date, end_date):
        """Chuyển tiếp yêu cầu lấy số liệu thống kê xuống service"""
        return self.service.get_product_stats_by_date(product_id, start_date, end_date)