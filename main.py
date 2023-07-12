from typing import Optional, Tuple, Union
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import customtkinter as ctk
from tkinter import ttk


class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self._set_appearance_mode("dark")

        self.geometry("1000x600")

        self.open_button = ctk.CTkButton(self, command=self.open_excel, text="Open excel")
        self.open_button.place(relx=0.5, rely=0.8, anchor="center")
        self.excel = Excel(self)

        
        # self.bind("<FocusOut>", exit)
        
        self.mainloop()

    def open_excel(self):
        path = ctk.filedialog.askopenfile(filetypes=[("Excel", "*.xlsx")])
        if path is not None:
            wb = load_workbook(path.name)
            ws = wb.worksheets[0]
            for row in range(1, ws.max_row + 1):
                for column in range(1, ws.max_column + 1):
                    pos = f"{get_column_letter(column)}{row}"
                    value = ws[pos].value
                    print(row, column)
                    self.excel.set(row, column=column, value=value)


class Excel(ttk.Treeview):
    def __init__(self, parent):
        super().__init__(parent, show="headings")

        nums = tuple(range(27))
        self["columns"] = nums

        self.column(0, stretch=False, width=50)

        for col in nums[1:]:
            self.column(col, stretch=False, width=100)
            self.heading(col, text=get_column_letter(col))
            self.insert("", index="end", iid=col)
            self.set(col, 0, col)
            
        self.bind("<MouseWheel>", lambda _: "break")
        self.place(x=0, y=0, relwidth=1, relheight=0.7)


MainApp()
