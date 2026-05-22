# services/remote_provider.py
from .base_provider import BaseProvider
import requests

class RemoteProvider(BaseProvider):
    def __init__(self, api_url):
        self.api_url = api_url
        
    def get_products(self):
        # response = requests.get(f"{self.api_url}/products")
        # return response.json()
        return "Dữ liệu từ API Server"

    def add_transaction(self, data):
        # response = requests.post(f"{self.api_url}/transactions", json=data)
        pass