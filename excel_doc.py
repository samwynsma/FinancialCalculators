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