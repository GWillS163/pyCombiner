"""
API Client service module
"""

import requests
from utils.helper import format_output

API_URL = "https://api.example.com/data"

def get_data():
    # Simulate an API call. In a real scenario, the following lines would be used.
    # response = requests.get(API_URL)
    # text_data = response.text
    text_data = "raw data from api"
    return format_output(text_data) 