import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class GetExcelInformation:
    def __init__(self):
        self.current_page = 0
    
    def create_gui(self):
        window = tk.Toplevel()
        self.window = window
        window.title("Financial Information Window")
        window.geometry("460x520")
        window.resizable(False, False)
        
        header = tk.Label(
            window,
            text="Display Document Information",
            font=("Segoe UI", 16, "bold"),
            wraplength=420,
            justify="center",
            pady=12,
        )
        header.pack()

        instructions = tk.Label(
            window,
            text="Select a page from the excel file and see what's on it.",
            font=("Segoe UI", 10),
            wraplength=420,
            justify="center",
        )
        instructions.pack(pady=(0, 10))

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)
        
        #tk.Button(button_frame, text="Export CSV", width=16, command=self.export_csv).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=1, padx=6)

        window.grab_set()
        window.mainloop()

def display_info():
    return