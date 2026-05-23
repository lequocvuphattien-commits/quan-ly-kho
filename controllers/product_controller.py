from services.data_service import DataService

class ProductController:
    def __init__(self):
        self.service = DataService()
        self.sheet = self.service.sheet_products

    def get_all_products(self):
        data = self.sheet.get_all_values()
        products = []
        # Bỏ qua dòng tiêu đề
        for row in data[1:]:
            if len(row) >= 5:
                # Dùng dict để tránh lỗi định nghĩa class lặp lại
                product = {
                    'id': row[0],
                    'code': str(row[1]).strip(),
                    'name': str(row[2]).strip(),
                    'unit': str(row[3]).strip(),
                    'stock': float(row[4]) if row[4] else 0.0
                }
                # Chuyển dict thành object để tương thích với logic cũ
                product_obj = type('Product', (object,), product)
                products.append(product_obj)
        return products

    def add_product(self, code, name, unit):
        if self.service.check_product_exists(code):
            return False, f"Mã '{code}' đã tồn tại!"
        self.service.add_product(code, name, unit)
        return True, "Thêm thành công!"