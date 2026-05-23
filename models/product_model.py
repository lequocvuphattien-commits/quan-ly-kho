class Product:
    def __init__(self, id=None, code=None, name=None, unit=None, stock=0.0):
        self.id = id
        self.code = str(code).strip() if code else ""
        self.name = str(name).strip() if name else ""
        self.unit = str(unit).strip() if unit else ""
        
        # Xử lý an toàn số liệu
        try:
            self.stock = float(stock) if str(stock).strip() != "" else 0.0
        except (ValueError, TypeError):
            self.stock = 0.0

    def to_dict(self):
        """Trả về dictionary thuần túy để Streamlit cache được"""
        return {
            "ID": self.id,
            "Mã": self.code,
            "Tên": self.name,
            "Đvt": self.unit,
            "Tồn": self.stock
        }

    @staticmethod
    def format_number(num):
        try:
            return "{:,.1f}".format(float(num))
        except (ValueError, TypeError):
            return "0.0"

    @staticmethod
    def from_tuple(data):
        # Đảm bảo dữ liệu đầu vào có đủ 5 phần tử để tránh lỗi IndexError
        if len(data) < 5:
            return Product(data[0] if len(data)>0 else None)
        return Product(id=data[0], code=data[1], name=data[2], unit=data[3], stock=data[4])