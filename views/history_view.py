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
        
        # Cấu trúc cột hiện tại của bạn: Ngày, Mã, Tên, Đvt, Loại, SL, Diễn Giải, Người, ID
        cols = ("Date", "Product", "Type", "Qty", "Note", "TransID")
        self.tree = ttk.Treeview(container, columns=cols, show="headings")
        self.tree.column("TransID", width=0, stretch=False)
        
        scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)
        self.tree.pack(side="top", fill="both", expand=True)
        scrollbar_x.pack(side="bottom", fill="x")
        
        # Tiêu đề hiển thị
        titles = ["Ngày", "Tên Hàng", "Loại", "Số Lượng", "Diễn Giải", "ID Giao Dịch"]
        for i, col in enumerate(cols):
            self.tree.heading(col, text=titles[i])
            self.tree.column(col, width=100, anchor="center")
        
        self.tree.tag_configure('nhap', foreground='green')
        self.tree.tag_configure('xuat', foreground='red')
        
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Hủy giao dịch này", command=self.undo_item)
        self.tree.bind("<Button-3>", self.show_menu)
        
        self.load_history()
        ttk.Button(self, text="Xuất Excel", command=self.export_excel).pack(pady=10)

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        history = self.t_controller.get_transaction_history()
        for row in history:
            # Kiểm tra dữ liệu: row[0]=Ngày, row[2]=Tên, row[4]=Loại, row[5]=SL, row[6]=Diễn Giải, row[8]=ID
            if len(row) < 9: continue
            
            t_type_raw = row[4].upper()
            tag = 'nhap' if 'NHẬP' in t_type_raw else 'xuat'
            t_type = "Nhập" if 'NHẬP' in t_type_raw else "Xuất"
            
            qty = Product.format_number(row[5]) # Cột Số lượng (index 5)
            note = row[6] # Cột Diễn Giải (index 6)
            trans_id = row[8] # Cột ID (index 8)
            
            self.tree.insert("", "end", values=(row[0], row[2], t_type, qty, note, trans_id), tags=(tag,))

    def export_excel(self):
        history = self.t_controller.get_transaction_history()
        if not history:
            messagebox.showwarning("Thông báo", "Không có dữ liệu!")
            return
            
        data = []
        for row in history:
            # Lấy đúng cột: Ngày, Tên, Loại, Số lượng, Diễn giải
            t_type = "Nhập" if 'NHẬP' in row[4].upper() else "Xuất"
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
            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:B', 30)
            worksheet.set_column('C:C', 10)
            worksheet.set_column('E:E', 30) # Cột Diễn Giải
            
            writer.close()
            messagebox.showinfo("Thông báo", "Xuất file thành công!")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

    # ... (giữ nguyên các hàm show_menu, undo_item, delete_selected_transaction)