from services.data_service import DataService
from models.product_model import Product

class ProductController:
    def __init__(self):
        self.service = DataService()
        self.sheet = self.service.sheet_products

    def get_all_products(self):
        data = self.service.get_products() # Lấy từ DataService
        products = []
        for row in data:
            # Kiểm tra xem dòng có đủ cột không (ít nhất 5 cột: ID, Mã, Tên, Đvt, Tồn)
            if len(row) < 5:
                continue
                
            try:
                # Ép kiểu an toàn: nếu cột 4 (Tồn) không phải số thì cho bằng 0.0
                stock_val = float(row[4]) if row[4] and str(row[4]).strip() != "" else 0.0
            except (ValueError, TypeError):
                stock_val = 0.0
                
            # Tạo object Product (giả sử cấu trúc model của bạn là thế này)
            products.append(Product(
                id=row[0],
                code=row[1],
                name=row[2],
                unit=row[3],
                stock=stock_val
            ))
        return products

    def add_product(self, code, name, unit):
        if self.service.check_product_exists(code):
            return False, f"Mã '{code}' đã tồn tại!"
        self.service.add_product(code, name, unit)
        return True, "Thêm thành công!"