import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from tkcalendar import DateEntry
from controllers.product_controller import ProductController
from controllers.transaction_controller import TransactionController
from models.product_model import Product

class ReportView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.p_controller = ProductController()
        self.t_controller = TransactionController()
        
        ttk.Label(self, text="BÁO CÁO TỒN KHO", font=("Arial", 14, "bold")).pack(pady=10)
        
        # --- BỘ LỌC NGÀY ---
        filter_frame = ttk.Frame(self)
        filter_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(filter_frame, text="Từ:").pack(side="left")
        self.cal_start = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd')
        self.cal_start.pack(side="left", padx=5)
        
        ttk.Label(filter_frame, text="Đến:").pack(side="left")
        self.cal_end = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd')
        self.cal_end.pack(side="left", padx=5)
        
        ttk.Button(filter_frame, text="Lọc báo cáo", command=self.load_report).pack(side="left", padx=10)
        
        # --- BẢNG DỮ LIỆU --- Bổ sung "Note" (Ghi chú)
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        cols = ("Code", "Name", "Unit", "Open", "In", "Out", "Close", "Note")
        self.tree = ttk.Treeview(container, columns=cols, show="headings")
        self.tree.pack(side="top", fill="both", expand=True)
        
        titles = ["Mã HH", "Tên Hàng", "Đvt", "Tồn Đầu", "Nhập", "Xuất", "Tồn Cuối", "Ghi chú"]
        for i, col in enumerate(cols):
            self.tree.heading(col, text=titles[i])
            self.tree.column(col, width=100, anchor="center")
            
        self.tree.tag_configure('low_stock', foreground='red')
        
        ttk.Button(self, text="Xuất Excel", command=self.export_excel).pack(pady=5)
        self.load_report()

    def load_report(self):
        start = self.cal_start.get_date().strftime("%Y-%m-%d 00:00:00")
        end = self.cal_end.get_date().strftime("%Y-%m-%d 23:59:59")
        
        for i in self.tree.get_children(): self.tree.delete(i)
        
        for p in self.p_controller.get_all_products():
            dau, nhap, xuat = self.t_controller.get_product_stats_by_date(p.id, start, end)
            cuoi = dau + nhap - xuat
            tags = ('low_stock',) if cuoi < 5 else ()
            
            # Lấy trường ghi chú một cách an toàn
            ghi_chu = getattr(p, 'note', '')
            
            self.tree.insert("", "end", values=(p.code, p.name, p.unit, 
                                               Product.format_number(dau), 
                                               Product.format_number(nhap), 
                                               Product.format_number(xuat), 
                                               Product.format_number(cuoi),
                                               ghi_chu), tags=tags)

    def export_excel(self):
        start = self.cal_start.get_date().strftime("%Y-%m-%d 00:00:00")
        end = self.cal_end.get_date().strftime("%Y-%m-%d 23:59:59")
        
        report_data = []
        for p in self.p_controller.get_all_products():
            dau, nhap, xuat = self.t_controller.get_product_stats_by_date(p.id, start, end)
            ghi_chu = getattr(p, 'note', '')
            report_data.append([p.code, p.name, p.unit, dau, nhap, xuat, dau + nhap - xuat, ghi_chu])
        
        if not report_data:
            messagebox.showwarning("Thông báo", "Không có dữ liệu!")
            return
            
        df = pd.DataFrame(report_data, columns=["Mã HH", "Tên Hàng", "Đvt", "Tồn đầu", "Nhập", "Xuất", "Tồn cuối", "Ghi chú"])
        
        file_name = "BaoCao_TonKho.xlsx"
        try:
            df.to_excel(file_name, index=False)
            messagebox.showinfo("Thành công", f"Đã xuất file '{file_name}'!")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file: {e}")