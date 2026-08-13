import tkinter as tk
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook

class ExcelPage:

    def __init__(self):
        self.columns = []

    def __init__(self, info):
        self.columns = info