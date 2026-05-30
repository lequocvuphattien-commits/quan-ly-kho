import sys
import os
import uuid
import pandas as pd
import datetime

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
            
            # 1. Làm sạch và GHI NHỚ thứ tự cột gốc của Google Sheets
            original_columns = [str(col).strip() for col in data[0]]
            df.columns = original_columns
            
            # 2. Xử lý cột Ngày
            if 'Ngày' in df.columns:
                df['Ngày'] = df['Ngày'].astype(str).str.strip()
                
                # Dịch sang kiểu thời gian để tính toán sắp xếp
                parsed_dates = pd.to_datetime(df['Ngày'], dayfirst=True, errors='coerce')
                
                # Thêm cột Sort_Date và tiến hành sắp xếp mới nhất lên trên
                df['Sort_Date'] = parsed_dates
                df = df.sort_values(by='Sort_Date', ascending=False)
                
                # Format lại thành chuỗi Text chuẩn để giao diện không bị None
                df['Ngày'] = parsed_dates.dt.strftime('%d/%m/%Y %H:%M:%S').fillna(df['Ngày'])
                
            # 3. ÉP KHÓA THỨ TỰ CỘT (Khắc phục triệt để lỗi hiển thị sai cột)
            df = df[original_columns]
            
            return df
        return pd.DataFrame()
    
    def add_transaction(self, p_code, p_name, qty, t_type, note, user_name, bo_phan=""):
        """
        Thêm giao dịch vào Google Sheets với 9 cột đầy đủ (Cột I là Bộ phận)
        """
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

            # 3. Tạo dòng dữ liệu mới (Đúng chuẩn 9 cột, thêm bo_phan vào cuối)
            tz_vn = datetime.timezone(datetime.timedelta(hours=7))
            now_str = datetime.datetime.now(tz_vn).strftime("%d/%m/%Y %H:%M:%S")
            new_row = [now_str, p_code, p_name, dvt, t_type, safe_qty, note, user_name, bo_phan]
            
            # 4. THỰC THI GHI VÀO GOOGLE SHEETS
            self.sheet_transactions.append_row(new_row)
            
            return True
        except Exception as e:
            print(f"Lỗi khi thêm giao dịch: {e}")
            return False

    def check_product_exists(self, product_code):
        products = self.get_products()
        return any(str(p[1]).strip().lower() == str(product_code).strip().lower() for p in products)

    def add_product(self, code, name, unit, group, min_level, ghi_chu=""):
        new_id = str(uuid.uuid4())[:8].upper()
        # Đẩy dữ liệu mới vào với 8 cột để khớp file Google Sheets
        self.sheet_products.append_row([new_id, code, name, unit, 0.0, group, min_level, ghi_chu])
        
        try:
            all_data = self.sheet_products.get_all_values()
            if len(all_data) > 2:
                product_rows = [row for row in all_data[1:] if any(str(cell).strip() for cell in row)]
                product_rows.sort(key=lambda x: str(x[2]).strip().lower() if len(x) > 2 else "")
                
                # BƯỚC BẢO VỆ: Đổi thành 8 để bảo toàn cả cột Ghi chú
                cleaned_rows = [row + [""] * (8 - len(row)) for row in product_rows]
                
                try:
                    self.sheet_products.update(values=cleaned_rows, range_name="A2") 
                except TypeError:
                    self.sheet_products.update("A2", cleaned_rows)
                    
        except Exception as e:
            print(f"⚠️ Lỗi tự động sắp xếp danh mục hàng hóa: {e}")
            
        return True

    def delete_product(self, product_id):
        data = self.sheet_products.get_all_values()
        for i, row in enumerate(data):
            if len(row) > 1 and str(row[1]).strip() == str(product_id).strip():
                self.sheet_products.delete_rows(i + 1)
                return True
        return False

    def get_products(self):
        """Lấy toàn bộ danh sách hàng hóa từ Google Sheets"""
        try:
            data = self.sheet_products.get_all_values()
            if len(data) > 1:
                results = []
                for row in data[1:]:
                    # Bù đủ 8 cột: ID, Mã, Tên, Đvt, Tồn, Nhóm, Mức tối thiểu, Ghi chú
                    if len(row) < 8:
                        row += [""] * (8 - len(row))
                    
                    # Nếu ô Mức tối thiểu trên Google Sheets bỏ trống, ép về 0
                    if str(row[6]).strip() == "":
                        row[6] = "0"
                        
                    results.append(row)
                return results
            return []
        except Exception as e:
            print(f"Lỗi khi tải danh sách hàng hóa: {e}")
            return []

    def update_product(self, product_code, new_name, new_unit, new_group, new_min_level, new_ghi_chu):
        try:
            data = self.sheet_products.get_all_values()
            for i, row in enumerate(data):
                if i > 0 and len(row) > 1 and str(row[1]).strip().upper() == str(product_code).strip().upper():
                    val_min = float(new_min_level)
                    self.sheet_products.update_cell(i + 1, 3, str(new_name))
                    self.sheet_products.update_cell(i + 1, 4, str(new_unit))
                    self.sheet_products.update_cell(i + 1, 6, str(new_group))
                    self.sheet_products.update_cell(i + 1, 7, val_min)
                    self.sheet_products.update_cell(i + 1, 8, new_ghi_chu)
                    return True
            return False
        except Exception as e:
            raise Exception(f"Lỗi ghi dữ liệu: {e}")
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
        kho_nhap = [str(r[0]).strip() for r in data[1:] if len(r) > 0 and str(r[0]).strip()]
        kho_xuat = [str(r[1]).strip() for r in data[1:] if len(r) > 1 and str(r[1]).strip()]
        bo_phan = [str(r[2]).strip() for r in data[1:] if len(r) > 2 and str(r[2]).strip()]
        
        return kho_nhap, kho_xuat, bo_phan

    def get_employees(self):
        """Lấy danh sách nhân viên từ sheet NhanVien, đảm bảo luôn trả về 5 cột"""
        data = self.sheet_employees.get_all_values()
        if len(data) > 1:
            results = []
            for row in data[1:]:
                row_data = row[:5]
                if len(row_data) < 5:
                    row_data += [""] * (5 - len(row_data))
                results.append(row_data)
            return results
        return []

    def check_employee_exists(self, emp_code):
        employees = self.get_employees()
        return any(str(emp[0]).strip().lower() == str(emp_code).strip().lower() for emp in employees if len(emp) > 0)

    def add_employee(self, emp_code, name, phone, role):
        self.sheet_employees.append_row([str(emp_code).upper(), str(name), str(phone), str(role), ""])  
        return True

    def update_employee(self, emp_code, new_name, new_phone, new_role):
        data = self.sheet_employees.get_all_values()
        for i, row in enumerate(data):
            if i > 0 and len(row) > 0 and str(row[0]).strip().upper() == str(emp_code).strip().upper():
                self.sheet_employees.update_cell(i + 1, 2, new_name)  
                self.sheet_employees.update_cell(i + 1, 3, new_phone) 
                self.sheet_employees.update_cell(i + 1, 4, new_role)  
                return True
        return False

    def delete_employee(self, emp_code):
        data = self.sheet_employees.get_all_values()
        for i, row in enumerate(data):
            if i > 0 and len(row) > 0 and str(row[0]).strip().upper() == str(emp_code).strip().upper():
                self.sheet_employees.delete_rows(i + 1)
                return True
        return False

    def check_login(self, username, password):
        employees = self.get_employees() 
        for emp in employees:
            if len(emp) >= 5:
                if str(emp[0]).strip().upper() == username.strip().upper() and str(emp[4]).strip() == password:
                    return {
                        "status": True, 
                        "name": emp[1], 
                        "role": str(emp[3]).strip()
                    }
        return {"status": False, "name": None, "role": None}

    def get_product_stats_by_date(self, product_id, start_date, end_date, history_data=None):
        """Tính toán tồn kho chuẩn xác bằng cách lùi thời gian từ Tồn Hiện Tại"""
        df = history_data if history_data is not None else self.get_history()
        
        # Lấy Tồn Kho Hiện Tại trực tiếp từ Danh mục hàng làm cột mốc
        current_stock = 0.0
        try:
            for p in self.get_products():
                if str(p[1]).strip().lower() == str(product_id).strip().lower():
                    current_stock = float(p[4]) if len(p) > 4 and str(p[4]).strip() else 0.0
                    break
        except Exception:
            pass

        if df is None or df.empty: 
            return current_stock, 0.0, 0.0
        
        try:
            df['Ngày'] = pd.to_datetime(df['Ngày'], dayfirst=True, errors='coerce')
            df['Mã HH'] = df['Mã HH'].astype(str).str.strip()
            df['Loại_Clean'] = df['Loại'].astype(str).str.strip().str.capitalize()
            
            qty_col = 'Số Lượng' if 'Số Lượng' in df.columns else 'Số lượng'
            df['Số lượng'] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
            
            target_id = str(product_id).strip()
            df_prod = df[df['Mã HH'] == target_id].copy()
            
            if df_prod.empty:
                return current_stock, 0.0, 0.0
                
            start = pd.to_datetime(start_date)
            end = pd.to_datetime(end_date)
            
            # --- [FIX BUG LOGIC KẾ TOÁN] ---
            # Tính ngược Tồn Đầu = Tồn Hiện Tại - Nhập (từ ngày báo cáo đến nay) + Xuất (từ ngày báo cáo đến nay)
            after_start_data = df_prod[df_prod['Ngày'] >= start]
            nhap_nguoc = after_start_data[after_start_data['Loại_Clean'] == 'Nhập']['Số lượng'].sum()
            xuat_nguoc = after_start_data[after_start_data['Loại_Clean'] == 'Xuất']['Số lượng'].sum()
            
            ton_dau = current_stock - nhap_nguoc + xuat_nguoc
            
            # Nhập / Xuất trong kỳ
            period_data = df_prod[(df_prod['Ngày'] >= start) & (df_prod['Ngày'] <= end)]
            nhap = period_data[period_data['Loại_Clean'] == 'Nhập']['Số lượng'].sum()
            xuat = period_data[period_data['Loại_Clean'] == 'Xuất']['Số lượng'].sum()
            
            return float(ton_dau), float(nhap), float(xuat)
        except Exception as e:
            print(f"Lỗi tính toán báo cáo tồn: {e}")
            return current_stock, 0.0, 0.0

    def delete_transaction(self, row_index, product_code, quantity, trans_type):
        self.sheet_transactions.delete_rows(row_index + 2) 
        change = -quantity if trans_type == "Nhập" else quantity
        self.update_stock(product_code, change, "Nhập")

    # =========================================================================
    # TÍNH NĂNG MỚI: XỬ LÝ LƯU HÀNG LOẠT SIÊU TỐC
    # =========================================================================
    def process_batch_transactions(self, transactions_list):
        """
        Ghi toàn bộ phiếu nhập/xuất nhiều dòng chỉ với 2 lần gọi API Google Sheets.
        Dữ liệu đầu vào (transactions_list): list các dict 
        [{'p_code': 'A', 'p_name':'B', 'qty': 10, 't_type': 'Nhập', 'note': '', 'user_name': '', 'bo_phan': ''}]
        """
        try:
            if not transactions_list:
                return False

            tz_vn = datetime.timezone(datetime.timedelta(hours=7))
            now_str = datetime.datetime.now(tz_vn).strftime("%d/%m/%Y %H:%M:%S")
            
            # 1. Tải danh mục SP 1 lần duy nhất để tra cứu ĐVT và sửa tồn kho trực tiếp trong RAM
            products_data = self.sheet_products.get_all_values()
            
            new_rows = []
            for txn in transactions_list:
                p_code = txn.get('p_code', '')
                try:
                    qty = float(txn.get('qty', 0))
                except:
                    qty = 0.0
                    
                t_type = txn.get('t_type', '')
                dvt = ""
                
                # Tìm và cập nhật Tồn kho (Cột 5 - index 4) trực tiếp trong mảng data
                for i, row in enumerate(products_data):
                    if i > 0 and len(row) > 1 and str(row[1]).strip().lower() == str(p_code).strip().lower():
                        dvt = str(row[3]).strip() if len(row) > 3 else ""
                        
                        current_stock = float(row[4]) if len(row) > 4 and str(row[4]).strip() != "" else 0.0
                        new_stock = current_stock + qty if t_type.strip().capitalize() == "Nhập" else current_stock - qty
                        
                        # Bù cột nếu mảng bị cụt
                        if len(row) < 5:
                            row += [""] * (5 - len(row))
                        products_data[i][4] = new_stock
                        break
                        
                # Đóng gói dữ liệu giao dịch
                new_row = [
                    now_str, p_code, txn.get('p_name', ''), dvt, t_type, 
                    qty, txn.get('note', ''), txn.get('user_name', ''), txn.get('bo_phan', '')
                ]
                new_rows.append(new_row)

            # 2. Bắn dữ liệu lên Server: Đúng 2 lệnh duy nhất cho 100 món hàng!
            if new_rows:
                self.sheet_transactions.append_rows(new_rows)
                
            self.sheet_products.update(values=products_data, range_name="A1")
            
            return True
        except Exception as e:
            print(f"Lỗi lưu hàng loạt: {e}")
            return False