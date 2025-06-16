import tkinter as tk
from tkinter import ttk, messagebox

class KeyColumnFrame(tk.Frame):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.shared = app.shared_data
        self.pack(padx=10, pady=10)
        self.create_widgets()

    def create_widgets(self):
        df1 = self.shared["df1"]
        df2 = self.shared["df2"]

        tk.Label(self, text="キー列（比較元側）:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(self, textvariable=self.shared["key_col1"], values=list(df1.columns), state="readonly").grid(row=0, column=1, sticky="w")

        tk.Label(self, text="キー列（比較先側）:").grid(row=1, column=0, sticky="w")
        ttk.Combobox(self, textvariable=self.shared["key_col2"], values=list(df2.columns), state="readonly").grid(row=1, column=1, sticky="w")

        tk.Button(self, text="戻る", command=self.app.show_file_select).grid(row=2, column=0, pady=10, sticky="w")
        tk.Button(self, text="終了", command=self.app.root.quit).grid(row=2, column=1, pady=10, sticky="e")
        tk.Button(self, text="次へ", command=self.app.show_column_mapping).grid(row=2, column=2, pady=10, sticky="w")