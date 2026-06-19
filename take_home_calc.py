import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class TakeHomeCalculator:
    def __init__(self, file_exists = False):
        self.pay_rate = -1.0
        self.percent_tax_rate = -1.0
        self.set_aside = -1.0
        self.pay_freq = "Hourly"
        self.period = "Month"
        self.file_exists = file_exists

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
        window.title("Retirement Duration Calculator")
        window.geometry("460x460")
        window.resizable(False, False)

        self.result_label = tk.Label(window, text="", font=("Segoe UI", 10), fg="green", wraplength=420, justify="center")
        self.result_label.pack(pady=(8, 0))

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Calculate", width=16, command=self.on_calculate).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=1, padx=6)

        window.grab_set()
        window.mainloop()
    
    def on_calculate(self):
        self.result_label.config(text="Results saved to InterestCalculation.xlsx")
        messagebox.showinfo("Take home document", "Take home calculation saved to InterestCalculation.xlsx.")

    

def take_home(file_exists):
    calculator = TakeHomeCalculator(file_exists)
    calculator.create_gui()
    return calculator.file_exists