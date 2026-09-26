import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

from excel_doc import ExcelDocument


class GetExcelInformation:
    def __init__(self, document=None):
        self.document = document if document is not None else ExcelDocument()
        self.current_page = self.document.current_page
        self.page_num = 0

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
        window.minsize(360, 360)

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

        navigation_frame = tk.Frame(window)
        navigation_frame.pack(pady=16)

        prev_button = tk.Button(
            navigation_frame,
            text="Prev",
            width=16,
            state=tk.DISABLED if self.current_page == 0 else tk.NORMAL,
        )
        prev_button.grid(row=0, column=0, padx=6)
        next_button = tk.Button(
            navigation_frame,
            text="Next",
            width=16,
            state=tk.DISABLED if self.current_page + 1 == self.document.pages else tk.NORMAL,
        )
        next_button.grid(row=0, column=1, padx=6)

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
                        prev_button.config(state=tk.DISABLED if index == 0 else tk.NORMAL)
                        break

            selected_page.trace_add("write", update_selected_page)
        else:
            empty_label = tk.Label(
                page_frame,
                text="No pages are available in the workbook yet.",
                font=("Segoe UI", 10),
            )
            empty_label.pack()

        content_frame = tk.Frame(window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        page_content = tk.Text(content_frame, wrap=tk.NONE, state=tk.DISABLED)
        vertical_scrollbar = tk.Scrollbar(
            content_frame, orient=tk.VERTICAL, command=page_content.yview
        )
        horizontal_scrollbar = tk.Scrollbar(
            content_frame, orient=tk.HORIZONTAL, command=page_content.xview
        )
        page_content.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )
        page_content.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        content_frame.grid_rowconfigure(0, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        def refresh_page_content():
            lines = []
            if self.document.page_list:
                columns = self.document.page_list[self.current_page].columns
                row_count = max((len(column) for column in columns), default=0)
                for row_index in range(row_count):
                    values = [
                        str(column[row_index]) if row_index < len(column) else ""
                        for column in columns
                    ]
                    lines.append(" | ".join(values))
            if not lines:
                lines.append("No information is available on this page.")

            page_content.configure(state=tk.NORMAL)
            page_content.delete("1.0", tk.END)
            page_content.insert("1.0", "\n".join(lines))
            page_content.configure(state=tk.DISABLED)

        def show_page(index):
            if not 0 <= index < len(self.document.page_list):
                return
            if selected_page.get() != page_names[index]:
                selected_page.set(page_names[index])
            self.current_page = index
            self.document.current_page = index
            current_page_label.config(text=f"Current page: {self.get_current_page_name()}")
            prev_button.config(state=tk.DISABLED if index == 0 else tk.NORMAL)
            next_button.config(
                state=tk.DISABLED if index == len(self.document.page_list) - 1 else tk.NORMAL
            )
            refresh_page_content()

        if self.document.page_list:
            def update_selected_page(*args):
                value = selected_page.get()
                for index, name in enumerate(page_names):
                    if name == value:
                        show_page(index)
                        break

            selected_page.trace_add("write", update_selected_page)
            prev_button.config(command=lambda: show_page(self.current_page - 1))
            next_button.config(command=lambda: show_page(self.current_page + 1))
            refresh_page_content()
        else:
            refresh_page_content()

        button_frame = tk.Frame(window)
        button_frame.pack(pady=16)

        tk.Button(button_frame, text="Export CSV", width=16, command=self.export_csv).grid(row=0, column=0, padx=6)
        tk.Button(button_frame, text="Export XLS", width=16, command=self.export_xls).grid(row=0, column=1, padx=6)
        tk.Button(button_frame, text="Quit", width=16, command=window.destroy).grid(row=0, column=2, padx=6)

        window.grab_set()
        window.mainloop()

    def export_csv(self):
        self.document.print_document_csv()

    def export_xls(self):
        self.document.print_document_xls()

    def next_page(self):
        return

    def prev_page(self):
        return


def display_info(document=None):
    info_window = GetExcelInformation(document)
    info_window.create_gui()
    return info_window