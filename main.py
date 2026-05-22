import tkinter as tk
from tkinter import ttk
from views.product_view import ProductView
from views.transaction_view import TransactionView
from views.report_view import ReportView
from views.history_view import HistoryView

class InventoryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PHẦN MỀM QUẢN LÝ KHO - GIAI ĐOẠN 2")
        # Mở rộng kích thước để hiển thị được hết các tab
        self.geometry("600x700") 

        # Style cho tab hiển thị rõ hơn
        style = ttk.Style()
        style.configure("TNotebook.Tab", padding=[10, 5])

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # Khởi tạo các View
        try:
            self.tab_product = ProductView(self.notebook)
            self.tab_transaction = TransactionView(self.notebook)
            self.tab_report = ReportView(self.notebook)
            self.tab_history = HistoryView(self.notebook)

            self.notebook.add(self.tab_product, text=" Danh mục ")
            self.notebook.add(self.tab_transaction, text=" Nhập/Xuất ")
            self.notebook.add(self.tab_report, text=" Báo cáo ")    
            self.notebook.add(self.tab_history, text=" Lịch sử ")
        except Exception as e:
            print(f"Lỗi khởi tạo View: {e}")

        # Lắng nghe sự kiện chuyển tab
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def on_tab_changed(self, event):
        selected_tab = self.notebook.select()
        tab_text = self.notebook.tab(selected_tab, "text").strip()
        
        # Gọi hàm làm mới dữ liệu của từng tab tương ứng
        if tab_text == "Nhập/Xuất":
            self.tab_transaction.refresh_products()
        elif tab_text == "Báo cáo":
            self.tab_report.load_report()
        elif tab_text == "Lịch sử":
            self.tab_history.load_history()

if __name__ == "__main__":
    app = InventoryApp()
    app.mainloop()