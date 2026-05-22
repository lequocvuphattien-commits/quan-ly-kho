class Product:
    def __init__(self, id=None, code=None, name=None, unit=None, stock=0.0):
        self.id = id
        self.code = code
        self.name = name
        self.unit = unit
        self.stock = float(stock)

    @staticmethod
    def format_number(num):
        """Chuyển số thành chuỗi dạng 5,000.0"""
        try:
            # : ,.1f sẽ hiển thị dấu phẩy phân cách và 1 chữ số thập phân
            return "{:,.1f}".format(float(num))
        except:
            return "0.0"

    @staticmethod
    def from_tuple(data):
        return Product(id=data[0], code=data[1], name=data[2], unit=data[3], stock=data[4])