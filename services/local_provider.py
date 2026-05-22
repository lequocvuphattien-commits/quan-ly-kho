import sqlite3
import os
from .base_provider import BaseProvider

class LocalProvider(BaseProvider):
    def __init__(self, db_name="inventory.db"):
        # Đảm bảo đường dẫn tới thư mục database
        self.db_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "database", db_name)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('CREATE TABLE IF NOT EXISTS products (id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT, unit TEXT, stock REAL)')
            conn.execute('CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY, product_id INTEGER, type TEXT, quantity REAL, note TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)')

    def execute_query(self, query, params=()):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            return cursor.lastrowid

    def fetch_data(self, query, params=()):
        with sqlite3.connect(self.db_path) as conn:
            return conn.execute(query, params).fetchall()

    # --- Các phương thức nghiệp vụ ---

    #def check_product_exists(self, code):
        """Kiểm tra mã hàng hóa đã tồn tại chưa"""
        res = self.fetch_data("SELECT id FROM products WHERE code = ?", (code,))
        return len(res) > 0

    def check_product_exists(self, code):
        """Kiểm tra mã hàng hóa không phân biệt chữ hoa/thường"""
        # Sử dụng LOWER() để so sánh cả hai vế ở dạng chữ thường
        query = "SELECT id FROM products WHERE LOWER(code) = LOWER(?)"
        res = self.fetch_data(query, (code,))
        return len(res) > 0

    def get_products(self):
        return self.fetch_data("SELECT * FROM products")

    def add_product(self, code, name, unit):
        self.execute_query("INSERT INTO products (code, name, unit, stock) VALUES (?, ?, ?, 0)", (code, name, unit))

    def delete_product(self, product_id):
        self.execute_query("DELETE FROM products WHERE id = ?", (product_id,))

    def update_product(self, product_id, name, unit):
        self.execute_query("UPDATE products SET name = ?, unit = ? WHERE id = ?", (name, unit, product_id))

    def add_transaction(self, product_id, quantity, transaction_type, note=""):
        # quantity ở vị trí thứ 2, transaction_type ở vị trí thứ 3
        change = quantity if transaction_type == 'IMPORT' else -quantity
        self.execute_query("UPDATE products SET stock = stock + ? WHERE id = ?", (change, product_id))
        self.execute_query("INSERT INTO transactions (product_id, type, quantity, note) VALUES (?, ?, ?, ?)", 
                           (product_id, transaction_type, quantity, note))

    def get_history(self):
        return self.fetch_data("SELECT t.created_at, p.name, t.type, t.quantity, t.note, t.id FROM transactions t JOIN products p ON t.product_id = p.id ORDER BY t.created_at DESC")

    def undo_transaction(self, trans_id):
        res = self.fetch_data("SELECT product_id, type, quantity FROM transactions WHERE id=?", (trans_id,))
        if res:
            p_id, t_type, qty = res[0]
            change = -qty if t_type == 'IMPORT' else qty
            self.execute_query("UPDATE products SET stock = stock + ? WHERE id = ?", (change, p_id))
            self.execute_query("DELETE FROM transactions WHERE id = ?", (trans_id,))

    def get_product_stats(self, product_id):
        # Tính toán tồn đầu + nhập - xuất
        nhap = self.fetch_data("SELECT IFNULL(SUM(quantity), 0) FROM transactions WHERE product_id=? AND type='IMPORT'", (product_id,))[0][0]
        xuat = self.fetch_data("SELECT IFNULL(SUM(quantity), 0) FROM transactions WHERE product_id=? AND type='EXPORT'", (product_id,))[0][0]
        stock = self.fetch_data("SELECT stock FROM products WHERE id=?", (product_id,))[0][0]
        return (stock - nhap + xuat, nhap, xuat)

    def get_product_stats_by_date(self, product_id, start, end):
        nhap = self.fetch_data("SELECT IFNULL(SUM(quantity), 0) FROM transactions WHERE product_id=? AND type='IMPORT' AND created_at BETWEEN ? AND ?", (product_id, start, end))[0][0]
        xuat = self.fetch_data("SELECT IFNULL(SUM(quantity), 0) FROM transactions WHERE product_id=? AND type='EXPORT' AND created_at BETWEEN ? AND ?", (product_id, start, end))[0][0]
        # Logic tính tồn kho trước ngày start
        before_nhap = self.fetch_data("SELECT IFNULL(SUM(quantity),0) FROM transactions WHERE product_id=? AND type='IMPORT' AND created_at < ?", (product_id, start))[0][0]
        before_xuat = self.fetch_data("SELECT IFNULL(SUM(quantity),0) FROM transactions WHERE product_id=? AND type='EXPORT' AND created_at < ?", (product_id, start))[0][0]
        before = before_nhap - before_xuat
        return (before, nhap, xuat)