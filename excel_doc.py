import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

from excel_page import ExcelPage

class ExcelDocument:
    def __init__(self):
        self.pages = 0
        self.page_list = []

    def create_page(self, info):
        page = ExcelPage(info)
        self.page_list.append(page)
        self.pages += 1

    def print_document_xls(self):
        workbook = Workbook()
        worksheet = workbook.active

        for page in self.page_list:
            for item in page.columns:
                for row in item:
                    print(row)

        workbook.save("FinancialDocuments.xlsx")
        print("XLS document created.")

    def print_document_word(self):
        return

    def print_document_csv(self):
        return


