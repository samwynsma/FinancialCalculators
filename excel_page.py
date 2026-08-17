import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class ExcelPage:

    def __init__(self):
        self.columns = []
        self.title = ""

    def __init__(self, info):
        self.columns = info

    def add_column(self, col):
        self.columns.append(col)

    def change_title(self, title):
        self.title = title