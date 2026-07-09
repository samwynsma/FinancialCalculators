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
        window.title("Net Worth Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        window.grab_set()
        window.mainloop()

    def get_percentile(self):
        percentile_values = [-76472.00, -45428.00, -26450.00, -14983.00, -9878.00, -4381.00, -831.80, 1.00, 182.20, 440.20, 
                            990.20, 2552.00, 4056.00, 5208.00, 6532.20, 7726.00, 9256.00, 10370.40, 11810.00, 13528.00,
                            15600.20, 18022.20, 20716.00, 23310.00, 27016.00, 30316.20, 34242.00, 39436.00, 44734.00, 51366.00,
                            57040.00, 62600.20, 67500.00, 73120.20, 79054.00, 84256.00, 89534.00, 96524.00, 101964.00, 110314.00,
                            117810.00, 125686.00, 132632.00, 141164.00, 147316.00, 155908.00, 164132.00, 172168.00, 181562.00, 192084.00,
                            202106.20, 212562.00]

def net_worth(file_exists):
    return file_exists