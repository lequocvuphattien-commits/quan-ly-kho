import tkinter as tk
from tkinter import ttk, messagebox
from controllers.product_controller import ProductController
from controllers.transaction_controller import TransactionController
from models.product_model import Product

class TransactionView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.p_controller = ProductController()
        self.t_controller = TransactionController()
        self.cart = []

        ttk.Label(self, text="NHẬP/XUẤT KHO", font=("Arial", 14, "bold")).pack(pady=10)

        ttk.Label(self, text="Tìm kiếm (Mã/Tên):").pack(anchor="w", padx=10)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_products) 
        ttk.Entry(self, textvariable=self.search_var).pack(fill="x", padx=10, pady=5)

        ttk.Label(self, text="Loại giao dịch:").pack(anchor="w", padx=10)
        self.combo_type = ttk.Combobox(self, values=["Nhập", "Xuất"], state="readonly")
        self.combo_type.pack(fill="x", padx=10, pady=5)
        self.combo_type.current(0)

        ttk.Label(self, text="Chọn hàng hóa:").pack(anchor="w", padx=10)
        self.combo_product = ttk.Combobox(self, state="readonly")
        self.combo_product.pack(fill="x", padx=10, pady=5)
        self.combo_product.bind("<<ComboboxSelected>>", self.on_product_select)

        info = ttk.Frame(self); info.pack(fill="x", padx=10, pady=5)
        ttk.Label(info, text="Tồn:").pack(side="left")
        self.lbl_stock = ttk.Label(info, text="0", font=("Arial", 10, "bold")); self.lbl_stock.pack(side="left", padx=10)
        ttk.Label(info, text="Đvt:").pack(side="left")
        self.lbl_unit = ttk.Label(info, text="-", font=("Arial", 10, "bold")); self.lbl_unit.pack(side="left", padx=10)

        ttk.Label(self, text="Số lượng:").pack(anchor="w", padx=10)
        self.entry_qty = ttk.Entry(self); self.entry_qty.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(self, text="Thêm Hàng Hóa", command=self.add_to_cart).pack(pady=5)

        # Tạo Treeview để hiển thị giỏ hàng
        self.tree = ttk.Treeview(self, columns=("Name", "Unit", "Qty"), show="headings")
        self.tree.heading("Name", text="Tên hàng hóa"); self.tree.column("Name", width=200)
        self.tree.heading("Unit", text="Đvt"); self.tree.column("Unit", width=60)
        self.tree.heading("Qty", text="Số lượng"); self.tree.column("Qty", width=100)

        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        ttk.Button(self, text="Thực hiện", command=self.process_all).pack(pady=10)
        self.refresh_products()

    def refresh_products(self):
        self.products = self.p_controller.get_all_products()
        self.combo_product['values'] = [f"{p.code} - {p.name}" for p in self.products]

    def filter_products(self, *args):
        query = self.search_var.get().lower()
        filtered = [f"{p.code} - {p.name}" for p in self.products if query in p.code.lower() or query in p.name.lower()]
        self.combo_product['values'] = filtered

    def on_product_select(self, event):
        val = self.combo_product.get()
        if not val: return
        code = val.split(" - ")[0]
        for p in self.products:
            if p.code == code:
                self.lbl_stock.config(text=Product.format_number(p.stock))
                self.lbl_unit.config(text=p.unit)
                break

    def add_to_cart(self):
        val = self.combo_product.get()
        if not val: return
        try:
            qty = float(self.entry_qty.get().replace(".", "").replace(",", "."))
            code = val.split(" - ")[0]
            p = next((p for p in self.products if p.code == code), None)
            if p and self.combo_type.get() == "Xuất" and qty > p.stock:
                messagebox.showerror("Lỗi", f"Tồn kho không đủ! (Còn {p.stock})")
                return
            if p:
                self.cart.append({"product": p, "qty": qty})
                self.tree.insert("", "end", values=(p.name, p.unit, Product.format_number(qty)))
                self.entry_qty.delete(0, tk.END)
        except ValueError: messagebox.showerror("Lỗi", "Số lượng sai!")

    def process_all(self):
        if not self.cart: return
        
        # Lấy loại giao dịch để truyền vào Controller
        t = "Nhập" if self.combo_type.get() == "Nhập" else "Xuất" 
        
        for i in self.cart: 
            p = i['product']
            qty = i['qty']
            
            # Gọi controller với đầy đủ tham số
            self.t_controller.process_transaction(
                product_id=p.id, 
                product_name=p.name,
                dvt=p.unit,
                trans_type=t, 
                quantity=qty, 
                note=""
            )
            
        messagebox.showinfo("Thông báo", "Đã xử lý!")
        self.cart = []
        [self.tree.delete(i) for i in self.tree.get_children()]

    def delete_transaction(self, trans_id):
        """Hủy một giao dịch và tự động hoàn tác tồn kho"""
        # Sửa lỗi gọi sai service, chuyển sang gọi thông qua t_controller
        self.t_controller.undo_transaction(trans_id)
        return True, "Đã hủy giao dịch thành công!"