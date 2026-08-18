import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class ExcelPage:

    def __init__(self, info=None):
        self.columns = info if info is not None else []
        self.title = ""

    def add_column(self, col):
        self.columns.append(col)

    def change_title(self, title):
        self.title = title