from abc import ABC, abstractmethod

class BaseProvider(ABC):
    @abstractmethod
    def get_products(self): pass
    @abstractmethod
    def add_product(self, code, name, unit): pass
    @abstractmethod
    def delete_product(self, product_id): pass
    @abstractmethod
    def update_product(self, product_id, name, unit): pass
    
    @abstractmethod
    def add_transaction(self, product_id, quantity, transaction_type, note): pass
    @abstractmethod
    def get_history(self): pass
    @abstractmethod
    def undo_transaction(self, trans_id): pass
    
    @abstractmethod
    def get_product_stats(self, product_id): pass
    @abstractmethod
    def get_product_stats_by_date(self, product_id, start, end): pass