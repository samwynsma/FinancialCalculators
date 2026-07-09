import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class NetWorthCalculator:
    def __init__(self, file_exists = False):
        self.file_exists = file_exists
        self.categories = []
        self.amount_per_category = []
        self.total = 0.0
        self.percentile = 0.0
    
    def parse_float(self, raw_value):
        if raw_value is None:
            raise ValueError("Missing value")
        if isinstance(raw_value, (int, float)):
            return float(raw_value)
        raw_value = raw_value.strip().replace(",", "")
        if raw_value.endswith("%"):
            raw_value = raw_value[:-1]
        if raw_value.startswith("$"):
            raw_value = raw_value[1:]
        return float(raw_value)
    
    def create_gui(self):
        window = tk.Toplevel()
        window.title("Food Cost Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        window.grab_set()
        window.mainloop()

    def get_percentile(self):
        percentile_values = []

def net_worth(file_exists):
    return file_exists