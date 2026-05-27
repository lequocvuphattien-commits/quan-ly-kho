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
        
        # Cột nội bộ: Ngày, Tên Hàng, Loại, SL, Diễn Giải
        cols = ("Date", "Product", "Type", "Qty", "Diễn Giải")
        self.tree = ttk.Treeview(container, columns=cols, show="headings")
        self.tree.column("TransID", width=0, stretch=False)
        
        scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)
        self.tree.pack(side="top", fill="both", expand=True)
        scrollbar_x.pack(side="bottom", fill="x")
        
        # Tiêu đề hiển thị (Đã đổi Ghi Chú thành Diễn Giải)
        titles = ["Ngày", "Tên Hàng", "Loại", "Số Lượng", "Diễn Giải"]
        for i, col in enumerate(cols):
            self.tree.heading(col, text=titles[i])
            self.tree.column(col, width=100, anchor="center")
            self.tree.column("Qty", anchor="w")
        
        self.tree.tag_configure('nhap', foreground='green')
        self.tree.tag_configure('xuat', foreground='red')
        
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Hủy giao dịch này", command=self.undo_item)
        self.tree.bind("<Button-3>", self.show_menu)
        
        self.load_history()
        ttk.Button(self, text="Xuất Excel", command=self.export_excel).pack(pady=10)

    def show_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu.post(event.x_root, event.y_root)

    def undo_item(self):
        if not self.tree.selection(): return
        item = self.tree.selection()[0]
        # TransID nằm ở cột cuối (index 5 trong values)
        trans_id = self.tree.item(item)['values'][5]
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn hủy giao dịch này?"):
            self.t_controller.undo_transaction(trans_id)
            self.load_history()
            messagebox.showinfo("Thông báo", "Đã hủy giao dịch!")

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        history = self.t_controller.get_transaction_history()
        for row in history:
            # Cấu trúc row từ Sheet (giả định): 
            # 0: Ngày, 1: Mã HH, 2: Tên HH, 3: Đvt, 4: Loại, 5: Số lượng, 6: Diễn giải, 7: Người, 8: ID
            if len(row) < 9: continue
            
            t_type_raw = str(row[4]).upper()
            tag = 'nhap' if 'NHẬP' in t_type_raw else 'xuat'
            t_type = "Nhập" if 'NHẬP' in t_type_raw else "Xuất"
            
            formatted_qty = Product.format_number(row[5]) # Lấy Số lượng (index 5)
            note = row[6]                                 # Lấy Diễn Giải (index 6)
            trans_id = row[8]                             # Lấy ID (index 8)
            
            self.tree.insert("", "end", values=(row[0], row[2], t_type, formatted_qty, note, trans_id), tags=(tag,))

    def delete_selected_transaction(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một giao dịch!")
            return
        
        trans_id = self.tree.item(selected_item)['values'][-1]
        
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn hủy giao dịch này?"):
            success, message = self.t_controller.delete_transaction(trans_id)
            if success:
                messagebox.showinfo("Thành công", message)
                self.load_history()

    def export_excel(self):
        history = self.t_controller.get_transaction_history()
        if not history:
            messagebox.showwarning("Thông báo", "Không có dữ liệu!")
            return
            
        data = []
        for row in history:
            # Lấy đúng các cột: 0: Ngày, 2: Tên Hàng, 4: Loại, 5: Số lượng, 6: Diễn Giải
            t_type = "Nhập" if 'NHẬP' in str(row[4]).upper() else "Xuất"
            data.append([row[0], row[2], t_type, row[5], row[6]])
            
        df = pd.DataFrame(data, columns=["Ngày", "Tên Hàng", "Loại", "Số Lượng", "Diễn Giải"])
        
        try:
            file_name = "LichSuGiaoDich.xlsx"
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='LichSu')
            
            workbook = writer.book
            worksheet = writer.sheets['LichSu']
            
            # Định dạng
            num_format = workbook.add_format({'num_format': '#,##0'})
            worksheet.set_column('D:D', 15, num_format)
            
            # Căn chỉnh cột
            worksheet.set_column('A:A', 20) # Ngày
            worksheet.set_column('B:B', 30) # Tên Hàng
            worksheet.set_column('C:C', 10) # Loại
            worksheet.set_column('D:D', 15) # Số Lượng
            worksheet.set_column('E:E', 30) # Diễn Giải
            
            writer.close()
            messagebox.showinfo("Thông báo", f"Đã xuất file '{file_name}' thành công!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file: {e}")