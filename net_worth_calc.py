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
                            202106.20, 212562.00, 223554.00, 238034.00, 250380.00, 261644.00, 274944.00, 288614.00, 298844.00, 312622.00,
                            327622.00, 347520.00, 366448.00, 384910.00, 402800.00, 415460.00, 429190.00, 447958.00, 468284.20, 493068.00,
                            521000.20, 551988.00, 587968.00, 622546.00, 685340.00, 697576.00, 743564.00, 785484.00, 836944.00, 891750.00,
                            947453.00, 1009860.00, 1078294.00, 1154634.00, 1234848.00, 1308426.00, 1399334.00, 1510942.00, 1693542.00, 1920758.00,
                            2157988.00, 2382960.00, 2692160.00, 3088722.00, 3779600.00, 4699180.20, 6150980.00, 8464740.20, 13666778.00]
        for i in range(len(percentile_values)):
            if percentile_values[i] > self.total:
                self.percentile = i
                return
        self.percentile = 99

def net_worth(file_exists):
    calculator = NetWorthCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists