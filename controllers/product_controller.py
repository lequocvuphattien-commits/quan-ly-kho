from services.data_service import DataService
from models.product_model import Product

class ProductController:
    def __init__(self):
        self.service = DataService()
        self.sheet = self.service.sheet_products

    def get_all_products(self):
        data = self.service.get_products()
        products = []
        for row in data:
            if len(row) < 5:
                continue
                
            try:
                stock_val = float(row[4]) if row[4] and str(row[4]).strip() != "" else 0.0
            except (ValueError, TypeError):
                stock_val = 0.0
                
            # Đọc cột Nhóm (Cột F - index 5). Nếu ô trống hoặc không có sẽ gán mặc định
            group_val = str(row[5]).strip() if len(row) > 5 and str(row[5]).strip() != "" else "Chưa phân nhóm"
                
            products.append(Product(
                id=row[0],
                code=row[1],
                name=row[2],
                unit=row[3],
                stock=stock_val,
                group=group_val  # Gán giá trị nhóm vào đối tượng
            ))
        return products

    def add_product(self, code, name, unit, group):
        if self.service.check_product_exists(code):
            return False, f"Mã '{code}' đã tồn tại!"
        self.service.add_product(code, name, unit)
        return True, "Thêm thành công!"