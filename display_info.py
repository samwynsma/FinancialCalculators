import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

from excel_doc import ExcelDocument


class GetExcelInformation:
    def __init__(self, document=None):
        self.document = document if document is not None else ExcelDocument()
        self.current_page = self.document.current_page

    def get_current_page_name(self):
        if not self.document.page_list:
            return "No page available"

        page = self.document.page_list[self.current_page]
        if getattr(page, "title", ""):
            return page.title
        return f"Page {self.current_page + 1}"

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

        current_page_label = tk.Label(
            window,
            text=f"Current page: {self.get_current_page_name()}",
            font=("Segoe UI", 11, "bold"),
            wraplength=420,
            justify="center",
            pady=8,
        )
        current_page_label.pack()

        page_frame = tk.Frame(window)
        page_frame.pack(pady=12)

        if self.document.page_list:
            page_names = [
                page.title if getattr(page, "title", "") else f"Page {index + 1}"
                for index, page in enumerate(self.document.page_list)
            ]
            selected_page = tk.StringVar(value=page_names[self.current_page])
            page_picker = tk.OptionMenu(page_frame, selected_page, *page_names)
            page_picker.pack()

            def update_selected_page(*args):
                value = selected_page.get()
                for index, name in enumerate(page_names):
                    if name == value:
                        self.current_page = index
                        self.document.current_page = index
                        current_page_label.config(text=f"Current page: {self.get_current_page_name()}")
                        break

            selected_page.trace_add("write", update_selected_page)
        else:
            empty_label = tk.Label(
                page_frame,
                text="No pages are available in the workbook yet.",
                font=("Segoe UI", 10),
            )
            empty_label.pack()

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Export CSV", width=16, command=self.export_csv).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=1, padx=6)

        window.grab_set()
        window.mainloop()

    def export_csv(self):
        return

    def export_xls(self):
        return


def display_info(document=None):
    info_window = GetExcelInformation(document)
    info_window.create_gui()
    return info_window