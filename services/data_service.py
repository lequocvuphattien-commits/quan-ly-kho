import sys
import os
import uuid
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from controllers.google_provider import GoogleProvider

class DataService:
    def __init__(self, mode="ONLINE"):
        self.provider = GoogleProvider()
        self.sheet_transactions = self.provider.get_sheet("Transactions")
        self.sheet_products = self.provider.get_sheet("Products")
        self.sheet_config = self.provider.get_sheet("Config")
        self.sheet_employees = self.provider.get_sheet("NhanVien")

    def get_history(self):
        data = self.sheet_transactions.get_all_values()
        if len(data) > 1:
            cleaned_data = [row[:7] for row in data[1:]]
            cleaned_data = [row + [""] * (7 - len(row)) for row in cleaned_data]
            df = pd.DataFrame(cleaned_data, columns=["date", "product_id", "product_name", "Đvt" "type", "qty", "note", "emp_name"])
            return df.values.tolist()
        return []

    def add_transaction(self, product_id, product_name, Đvt, qty, trans_type, note, emp_name=""):
        date_str = pd.Timestamp.now(tz='Asia/Ho_Chi_Minh').strftime("%Y-%m-%d %H:%M:%S")
        self.sheet_transactions.append_row([date_str, str(product_id), str(product_name), str(Đvt), trans_type.upper(), float(qty), str(note), str(emp_name)])

    def get_products(self):
        data = self.sheet_products.get_all_values()
        return data[1:] if len(data) > 1 else []

    def check_product_exists(self, product_code):
        products = self.get_products()
        return any(str(p[1]).strip().lower() == str(product_code).strip().lower() for p in products)

    def add_product(self, code, name, unit):
        new_id = str(uuid.uuid4())[:8].upper()
        self.sheet_products.append_row([new_id, code, name, unit, 0.0])
        return True

    def delete_product(self, product_id):
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id:
                self.sheet_products.delete_rows(i + 1)
                return True
        return False

    def update_product(self, product_id, new_name, new_unit):
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id:
                self.sheet_products.update_cell(i + 1, 3, new_name)
                self.sheet_products.update_cell(i + 1, 4, new_unit)
                return True
        return False

    def update_stock(self, product_code, qty, trans_type):
        records = self.sheet_products.get_all_values()
        for i, row in enumerate(records):
            if i > 0 and len(row) > 1 and str(row[1]).strip().lower() == str(product_code).strip().lower():
                current_stock = float(row[4]) if len(row) > 4 and row[4] else 0.0
                new_stock = current_stock + float(qty) if trans_type.strip().capitalize() == "Nhập" else current_stock - float(qty)
                self.sheet_products.update_cell(i + 1, 5, new_stock)
                return True
        return False

    def get_config_options(self):
        data = self.sheet_config.get_all_values()
        return ([str(r[0]) for r in data[1:] if r[0]], [str(r[1]) for r in data[1:] if len(r)>1 and r[1]])

    def get_employees(self):
        """Lấy danh sách nhân viên từ sheet NhanVien, đảm bảo luôn trả về 5 cột"""
        data = self.sheet_employees.get_all_values()
        if len(data) > 1:
            results = []
            for row in data[1:]: # Bỏ qua dòng tiêu đề
                # Lấy 5 cột đầu tiên
                row_data = row[:5]
                # Nếu dòng thiếu cột, bù bằng chuỗi rỗng
                if len(row_data) < 5:
                    row_data += [""] * (5 - len(row_data))
                results.append(row_data)
            return results
        return []

    def check_employee_exists(self, emp_code):
        """Kiểm tra mã nhân viên đã tồn tại chưa"""
        employees = self.get_employees()
        return any(str(emp[0]).strip().lower() == str(emp_code).strip().lower() for emp in employees if len(emp) > 0)

    def add_employee(self, emp_code, name, phone, role):
        """Thêm nhân viên mới"""
        # Cấu trúc: [Mã NV, Tên NV, Số điện thoại, Chức vụ, Mật khẩu]
        self.sheet_employees.append_row([str(emp_code).upper(), str(name), str(phone), str(role), ""])  # Mật khẩu mặc định là chuỗi rỗng
        return True

    def update_employee(self, emp_code, new_name, new_phone, new_role):
        """Cập nhật thông tin nhân viên (Sửa trực tiếp trên lưới)"""
        data = self.sheet_employees.get_all_values()
        for i, row in enumerate(data):
            if i > 0 and len(row) > 0 and str(row[0]).strip().upper() == str(emp_code).strip().upper():
                self.sheet_employees.update_cell(i + 1, 2, new_name)  # Cột B: Tên NV
                self.sheet_employees.update_cell(i + 1, 3, new_phone) # Cột C: SĐT
                self.sheet_employees.update_cell(i + 1, 4, new_role)  # Cột D: Chức vụ
                return True
        return False

    def delete_employee(self, emp_code):
        """Xóa nhân viên"""
        data = self.sheet_employees.get_all_values()
        for i, row in enumerate(data):
            if i > 0 and len(row) > 0 and str(row[0]).strip().upper() == str(emp_code).strip().upper():
                self.sheet_employees.delete_rows(i + 1)
                return True
        return False
    
    # Hàm kiểm tra đăng nhập dựa trên Mã NV (Username) và Mật khẩu (giả sử cột 4 là mật khẩu)
        return [row[:5] + [""] * (5 - len(row)) for row in data[1:]] if len(data) > 1 else []

    def check_login(self, username, password):
        """Kiểm tra đăng nhập và trả về Tên + Chức vụ"""
        employees = self.get_employees() 
        for emp in employees:
            # Cấu trúc nhân viên: [Mã NV, Tên NV, SĐT, Chức vụ, Mật khẩu]
            # Index:                 0        1       2       3         4
            if len(emp) >= 5:
                if str(emp[0]).strip().upper() == username.strip().upper() and str(emp[4]).strip() == password:
                    return {
                        "status": True, 
                        "name": emp[1], 
                        "role": str(emp[3]).strip() # Phải có dòng này để trả về 'role'
                    }
        return {"status": False, "name": None, "role": None}
    
    def delete_transaction(self, row_index, product_code, quantity, trans_type):
        """
        Xóa giao dịch và hoàn tác số lượng tồn kho
        """
        # 1. Xóa dòng trong sheet Transactions
        # Lưu ý: row_index lấy từ st.data_editor thường là index của DataFrame, 
        # nên khi map vào Google Sheet phải cộng thêm 2 (1 cho header, 1 vì index bắt đầu từ 0)
        self.sheet_transactions.delete_rows(row_index + 2) 
        
        # 2. Hoàn tác tồn kho
        # Nếu đã nhập kho (Nhập), giờ xóa đi thì phải trừ tồn (tồn giảm)
        # Nếu đã xuất kho (Xuất), giờ xóa đi thì phải cộng lại tồn (tồn tăng)
        change = -quantity if trans_type == "Nhập" else quantity
        
        # Gọi lại hàm update_stock đã có sẵn của bạn để đồng bộ
        self.update_stock(product_code, change, "Nhập")
