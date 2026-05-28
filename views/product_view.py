from email.mime import message
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from controllers.product_controller import ProductController
from models.product_model import Product
from tkinter import messagebox

class ProductView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.controller = ProductController()
        
        # 1. Khung nhập liệu
        input_frame = ttk.LabelFrame(self, text="Thêm hàng hóa mới", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        # Hàng 1: Các ô nhập liệu
        ttk.Label(input_frame, text="Mã:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_code = ttk.Entry(input_frame, width=10)
        self.entry_code.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Tên:").grid(row=0, column=2, padx=5, pady=5)
        self.entry_name = ttk.Entry(input_frame, width=20)
        self.entry_name.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Đvt:").grid(row=0, column=4, padx=5, pady=5)
        self.entry_unit = ttk.Entry(input_frame, width=8)
        self.entry_unit.grid(row=0, column=5, padx=5, pady=5)

        # Hàng 2: Nút Thêm Hàng nằm ở giữa
        # columnspan=6 cho phép nút trải dài hết chiều rộng của 6 cột phía trên, 
        # sau đó dùng sticky="ew" để căn giữa hoặc điều khiển vị trí
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=1, column=0, columnspan=6, pady=10)
        
        ttk.Button(btn_frame, text="Thêm Hàng", command=self.add_item).pack()

        def handle_add_product(self):
            code = self.entry_code.get()
            name = self.entry_name.get()
            unit = self.entry_unit.get()
            # ... lấy các giá trị khác

            success, message = self.controller.add_product(code, name, unit)

            if success:
                 messagebox.showinfo("Thông báo", message)
            # Refresh bảng dữ liệu
            else:
                messagebox.showwarning("Lỗi trùng mã", message)
        
        # 2. Tiêu đề bảng
        ttk.Label(self, text="DANH MỤC HÀNG HÓA", font=("Arial", 12, "bold")).pack(pady=5)
        
        # 3. Treeview hiển thị
        self.tree = ttk.Treeview(self, columns=("ID", "Code", "Name", "Unit", "Stock", "Group"), show="headings")
        self.tree.heading("ID", text="ID"); self.tree.column("ID", width=30)
        self.tree.heading("Code", text="Mã"); self.tree.column("Code", width=80)
        self.tree.heading("Name", text="Tên hàng"); self.tree.column("Name", width=200)
        self.tree.heading("Unit", text="Đvt"); self.tree.column("Unit", width=60)
        self.tree.heading("Stock", text="Tồn"); self.tree.column("Stock", width=60)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.heading("Group", text="Nhóm"); self.tree.column("Group", width=100)
        
        # 4. Nút Xuất Excel
        ttk.Button(self, text="Xuất Excel", command=self.export_excel).pack(pady=5)
        
        # Menu chuột phải
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Sửa thông tin", command=self.edit_item)
        self.menu.add_command(label="Xóa hàng hóa", command=self.delete_item)
        self.tree.bind("<Button-3>", self.show_menu)

        self.load_data()

    def add_item(self):
        code = self.entry_code.get()
        name = self.entry_name.get()
        unit = self.entry_unit.get()
        group = self.combo_group.get() # Lấy giá trị nhóm

        # Gọi Controller với 4 tham số
        success, message = self.controller.add_product(code, name, unit, group)

        if success:
            messagebox.showinfo("Thông báo", message)
            # Clear ô nhập liệu
            self.entry_code.delete(0, 'end')
            self.entry_name.delete(0, 'end')
            self.entry_unit.delete(0, 'end')
            self.refresh_table() 
        else:
            messagebox.showwarning("Cảnh báo", message)
            
    def export_excel(self):
        products = self.controller.get_all_products()
        data = [[p.code, p.name, p.unit, p.stock, p.group] for p in products]
        
        if not data:
            messagebox.showwarning("Thông báo", "Không có dữ liệu!")
            return
            
        df = pd.DataFrame(data, columns=["Mã", "Tên Hàng", "Đvt", "Tồn", "Nhóm"])
        
        try:
            file_name = "DanhMucHangHoa.xlsx"
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Danh mục')
            
            # Lấy đối tượng workbook và worksheet để định dạng
            workbook = writer.book
            worksheet = writer.sheets['Danh mục']
            
            # Định dạng cho cột Tồn (cột D trong Excel, index là 3)
            # "#,##0" là định dạng có dấu phẩy hàng nghìn
            format_num = workbook.add_format({'num_format': '#,##0'})
            worksheet.set_column('D:D', 15, format_num)
            
            writer.close()
            messagebox.showinfo("Thông báo", f"Đã xuất file {file_name} thành công!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file: {e}")

    def show_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu.post(event.x_root, event.y_root)

    def delete_item(self):
        if not self.tree.selection(): return
        item = self.tree.selection()[0]
        data = self.tree.item(item)['values']
        if messagebox.askyesno("Xác nhận", f"Xóa hàng hóa '{data[2]}'?"):
            self.controller.delete_product(data[0])
            self.load_data()

    def edit_item(self):
        if not self.tree.selection(): return
        item = self.tree.selection()[0]
        data = self.tree.item(item)['values']
        edit_win = tk.Toplevel(self)
        edit_win.title("Sửa hàng hóa"); edit_win.geometry("300x200")
        ttk.Label(edit_win, text="Tên hàng:").pack(pady=5)
        entry_name = ttk.Entry(edit_win); entry_name.insert(0, data[2]); entry_name.pack()
        ttk.Label(edit_win, text="Đvt:").pack(pady=5)
        entry_unit = ttk.Entry(edit_win); entry_unit.insert(0, data[3]); entry_unit.pack()
        def save():
            self.controller.update_product(data[0], entry_name.get(), entry_unit.get())
            edit_win.destroy(); self.load_data()
        ttk.Button(edit_win, text="Lưu", command=save).pack(pady=15)

    def load_data(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for p in self.controller.get_all_products():
            self.tree.insert("", "end", values=(p.id, p.code, p.name, p.unit, Product.format_number(p.stock)))

    def refresh_table(self):
        """Xóa dữ liệu cũ và load lại danh sách sản phẩm mới từ controller"""
        # 1. Xóa toàn bộ dữ liệu đang hiển thị trong Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 2. Lấy dữ liệu mới từ controller
        products = self.controller.get_all_products()
        
        # 3. Chèn dữ liệu vào lại Treeview
        for p in products:
            self.tree.insert("", "end", values=(p.id, p.code, p.name, p.unit, p.stock, p.group))