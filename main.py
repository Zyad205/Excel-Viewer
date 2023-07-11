from typing import Optional, Tuple, Union
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import customtkinter as ctk
from tkinter import ttk


class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self._set_appearance_mode("dark")

        Excel(self)

        self.mainloop()


class Excel(ttk.Treeview):
    def __init__(self, parent):
        super().__init__(parent, show="headings")

        nums = tuple(range(27))
        self["columns"] = nums

        for col in nums[1:]:
            self.column(col, stretch=False)
            self.heading(col, text=get_column_letter(col))
        self.place(x=0, y=0, relwidth=1, relheight=0.7)
                

MainApp()