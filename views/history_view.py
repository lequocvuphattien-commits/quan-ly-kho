import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from controllers.transaction_controller import TransactionController
from models.product_model import Product

class HistoryView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.t_controller = TransactionController()
        
        ttk.Label(self, text="LỊCH SỬ GIAO DỊCH", font=("Arial", 14, "bold")).pack(pady=10)
        
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=5)
        
        cols = ("Date", "Product", "Type", "Qty", "Note", "TransID")
        self.tree = ttk.Treeview(container, columns=cols, show="headings")
        self.tree.column("TransID", width=0, stretch=False)
        
        scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)
        self.tree.pack(side="top", fill="both", expand=True)
        scrollbar_x.pack(side="bottom", fill="x")
        
        titles = ["Ngày", "Tên Hàng", "Loại", "Số Lượng", "Ghi Chú", "ID Giao Dịch"]
        for i, col in enumerate(cols):
            self.tree.heading(col, text=titles[i])
            self.tree.column(col, width=100, anchor="center")
        
        self.tree.tag_configure('nhap', foreground='green')
        self.tree.tag_configure('xuat', foreground='red')
        
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Hủy giao dịch này", command=self.undo_item)
        self.tree.bind("<Button-3>", self.show_menu)
        
        self.load_history()

        # Thêm đoạn này vào trong __init__ của HistoryView
        ttk.Button(self, text="Xuất Excel", command=self.export_excel).pack(pady=10)

    def show_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu.post(event.x_root, event.y_root)

    def undo_item(self):
        if not self.tree.selection(): return
        item = self.tree.selection()[0]
        trans_id = self.tree.item(item)['values'][5]
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn hủy giao dịch này?"):
            self.t_controller.undo_transaction(trans_id)
            self.load_history()
            messagebox.showinfo("Thông báo", "Đã hủy giao dịch!")

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        history = self.t_controller.get_transaction_history()
        for row in history:
            if len(row) < 6: continue
            
            tag = 'nhap' if row[2] == 'IMPORT' else 'xuat'
            t_type = "Nhập" if row[2] == 'IMPORT' else "Xuất"
            
            # Sử dụng Product.format_number cho row[3] (cột số lượng)
            formatted_qty = Product.format_number(row[3])
            
            self.tree.insert("", "end", values=(row[0], row[1], t_type, formatted_qty, row[4], row[5]), tags=(tag,))

    def delete_selected_transaction(self):
        # 1. Lấy ID của dòng đang chọn trong Treeview
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một giao dịch để xóa!")
            return
        
        # 2. Lấy trans_id (giả sử cột cuối cùng của Treeview là trans_id)
        trans_id = self.tree.item(selected_item)['values'][-1]
        
        # 3. Xác nhận xóa
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn hủy giao dịch này? Tồn kho sẽ được cập nhật lại."):
            success, message = self.t_controller.delete_transaction(trans_id)
            if success:
                messagebox.showinfo("Thành công", message)
                self.load_history() # Tải lại bảng lịch sử

    def export_excel(self):
        # 1. Lấy dữ liệu lịch sử từ controller
        history = self.t_controller.get_transaction_history()
        
        if not history:
            messagebox.showwarning("Thông báo", "Không có dữ liệu giao dịch!")
            return
            
        # 2. Chuyển đổi dữ liệu để xuất (để nguyên số cho Excel xử lý định dạng)
        data = []
        for row in history:
            # row: [date, product_name, type, quantity, note, id]
            t_type = "Nhập" if row[2] == 'IMPORT' else "Xuất"
            data.append([row[0], row[1], t_type, row[3], row[4]])
            
        df = pd.DataFrame(data, columns=["Ngày", "Tên Hàng", "Loại", "Số Lượng", "Ghi Chú"])
        
        # 3. Xuất file bằng xlsxwriter
        try:
            file_name = "LichSuGiaoDich.xlsx"
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='LichSu')
            
            workbook = writer.book
            worksheet = writer.sheets['LichSu']
            
            # 4. Định dạng số có dấu phẩy (Cột D - index 3)
            num_format = workbook.add_format({'num_format': '#,##0'})
            worksheet.set_column('D:D', 15, num_format)
            
            # Căn chỉnh cột cho đẹp
            worksheet.set_column('A:A', 20) # Ngày
            worksheet.set_column('B:B', 30) # Tên Hàng
            worksheet.set_column('C:C', 10) # Loại
            worksheet.set_column('E:E', 20) # Ghi chú
            
            writer.close()
            messagebox.showinfo("Thông báo", f"Đã xuất file '{file_name}' thành công!")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file: {e}")

    