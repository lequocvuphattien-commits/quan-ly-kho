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
            # Tạo DataFrame từ dữ liệu
            df = pd.DataFrame(data[1:], columns=data[0])
            
            # 1. Làm sạch tên cột (tránh khoảng trắng thừa)
            df.columns = [str(col).strip() for col in df.columns]
            
            # 2. Ép kiểu cột "Ngày" (hoặc tên cột tương ứng trong sheet của bạn)
            # Giả sử cột ngày có tên là 'Ngày'
            if 'Ngày' in df.columns:
                df['Ngày'] = pd.to_datetime(df['Ngày'], dayfirst=True, errors='coerce')
                # Sắp xếp mới nhất lên trên
                df = df.sort_values(by='Ngày', ascending=False)
            
            return df
        return pd.DataFrame()
    
    # Đã đồng bộ tên biến dvt (viết thường)
    def add_transaction(self, p_code, p_name, qty, t_type, note, user_name):
        """
        Thêm giao dịch vào Google Sheets với 8 cột đầy đủ
        """
        import datetime # Khai báo để lấy giờ hệ thống
        try:
            # 1. Ép kiểu an toàn
            try:
                safe_qty = float(qty)
            except (ValueError, TypeError):
                safe_qty = 0.0
                
            if safe_qty <= 0:
                return False 

            # 2. Tự động tìm Đơn vị tính (Đvt) từ Sheet Products
            dvt = ""
            products = self.get_products()
            for p in products:
                if str(p[1]).strip().lower() == str(p_code).strip().lower():
                    dvt = str(p[3]).strip() if len(p) > 3 else ""
                    break

            # 3. Tạo dòng dữ liệu mới (Đúng chuẩn 8 cột)
            # Cấu trúc: [Ngày, Mã HH, Tên hàng hóa, Đvt, Loại, Số lượng, Diễn giải, Nhân Viên]
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_row = [now_str, p_code, p_name, dvt, t_type, safe_qty, note, user_name]
            
            # 4. THỰC THI GHI VÀO GOOGLE SHEETS (Đã mở khóa)
            self.sheet_transactions.append_row(new_row)
            
            return True
        except Exception as e:
            print(f"Lỗi khi thêm giao dịch: {e}")
            return False

    def get_products(self):
        data = self.sheet_products.get_all_values()
        return data[1:] if len(data) > 1 else []

    def check_product_exists(self, product_code):
        products = self.get_products()
        return any(str(p[1]).strip().lower() == str(product_code).strip().lower() for p in products)

    def add_product(self, code, name, unit):
        # 1. Tạo và thêm hàng hóa mới vào dòng cuối trước
        new_id = str(uuid.uuid4())[:8].upper()
        self.sheet_products.append_row([new_id, code, name, unit, 0.0])
        
        # 2. Tự động sắp xếp lại Sheet theo Tên hàng hóa (A-Z)
        try:
            all_data = self.sheet_products.get_all_values()
            if len(all_data) > 2:
                # BƯỚC BẢO VỆ 1: Lọc bỏ toàn bộ các dòng trống hoàn toàn ở cuối file
                # Tránh việc dòng trống bị mang đi sắp xếp và nhảy lên đầu sheet
                product_rows = [row for row in all_data[1:] if any(str(cell).strip() for cell in row)]
                
                # BƯỚC BẢO VỆ 2: Sắp xếp tăng dần theo Tên hàng hóa (Cột C -> index 2)
                product_rows.sort(key=lambda x: str(x[2]).strip().lower() if len(x) > 2 else "")
                
                # BƯỚC BẢO VỆ 3: Chuẩn hóa độ dài các dòng (Bù khoảng trống)
                # Google Sheets API sẽ báo lỗi nếu mảng có dòng dài dòng ngắn
                cleaned_rows = [row + [""] * (5 - len(row)) for row in product_rows]
                
                # BƯỚC BẢO VỆ 4: Cập nhật dữ liệu (Tương thích mọi phiên bản gspread)
                try:
                    # Dành cho gspread phiên bản mới (v6.0 trở lên)
                    self.sheet_products.update(values=cleaned_rows, range_name="A2") 
                except TypeError:
                    # Dành cho gspread phiên bản cũ (v5.x trở xuống)
                    self.sheet_products.update("A2", cleaned_rows)
                    
        except Exception as e:
            print(f"⚠️ Lỗi tự động sắp xếp danh mục hàng hóa: {e}")
            
        return True

    def delete_product(self, product_id):
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if row[1] == product_id:
                self.sheet_transactions.delete_rows(i + 1)
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

    # ==========================================
    # 🔥 ĐÃ SỬA LỖI ĐỌC TÊN CỘT THEO ẢNH BẠN CUNG CẤP 🔥
    # ==========================================
    def get_product_stats_by_date(self, product_id, start_date, end_date):
        df = self.get_history()
        if df.empty: return 0.0, 0.0, 0.0
        
        # Sửa thành "Ngày" thay vì "Ngày"
        # Sửa thành "Số lượng" thay vì "Số Lượng"
        try:
            df['Ngày'] = pd.to_datetime(df['Ngày'], errors='coerce')
            df['Mã HH'] = df['Mã HH'].astype(str).str.strip()
            df['Số lượng'] = pd.to_numeric(df['Số lượng'], errors='coerce').fillna(0)
            
            # Lọc
            target_id = str(product_id).strip()
            df_prod = df[df['Mã HH'] == target_id].copy()
            
            start = pd.to_datetime(start_date)
            past_data = df_prod[df_prod['Ngày'] < start]
            
            # Tính toán (IMPORT/EXPORT khớp với cột 'Loại' trong Sheet)
            ton_dau = (past_data[past_data['Loại'] == 'Nhập']['Số lượng'].sum() - 
                       past_data[past_data['Loại'] == 'Xuất']['Số lượng'].sum())
            
            period_data = df_prod[(df_prod['Ngày'] >= start) & (df_prod['Ngày'] <= pd.to_datetime(end_date))]
            nhap = period_data[period_data['Loại'] == 'Nhập']['Số lượng'].sum()
            xuat = period_data[period_data['Loại'] == 'Xuất']['Số lượng'].sum()
            
            return float(ton_dau), float(nhap), float(xuat)
        except Exception as e:
            print(f"Lỗi tính toán báo cáo tồn: {e}")
            return 0.0, 0.0, 0.0

    def delete_transaction(self, row_index, product_code, quantity, trans_type):
        """
        Xóa giao dịch và hoàn tác số lượng tồn kho
        """
        self.sheet_transactions.delete_rows(row_index + 2) 
        change = -quantity if trans_type == "Nhập" else quantity
        self.update_stock(product_code, change, "Nhập")