import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd

class FileSelectFrame(tk.Frame):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.shared = app.shared_data
        self.pack(padx=10, pady=10)
        self.create_widgets()

    def create_widgets(self):
        # 初期値を設定
        self.shared["start_row1"].set(1)
        self.shared["start_row2"].set(1)

        tk.Label(self, text="比較元ファイルパス:").grid(row=0, column=0, sticky="w")
        tk.Entry(self, textvariable=self.shared["file1_path"], width=50).grid(row=0, column=1, sticky="w")
        tk.Button(self, text="参照", command=lambda: self.select_file("file1_path")).grid(row=0, column=2)

        tk.Label(self, text="ヘッダー行:").grid(row=1, column=0, sticky="w")
        tk.Entry(self, textvariable=self.shared["start_row1"], width=5).grid(row=1, column=1, sticky="w")

        tk.Label(self, text="比較先ファイルパス:").grid(row=2, column=0, sticky="w")
        tk.Entry(self, textvariable=self.shared["file2_path"], width=50).grid(row=2, column=1, sticky="w")
        tk.Button(self, text="参照", command=lambda: self.select_file("file2_path")).grid(row=2, column=2)

        tk.Label(self, text="ヘッダー行:").grid(row=3, column=0, sticky="w")
        tk.Entry(self, textvariable=self.shared["start_row2"], width=5).grid(row=3, column=1, sticky="w")

        tk.Button(self, text="終了", command=self.app.root.quit).grid(row=4, column=1, sticky="e", pady=10)
        tk.Button(self, text="次へ", command=self.load_files).grid(row=4, column=2, sticky="w")

    def select_file(self, var_name):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
        if path:
            self.shared[var_name].set(path)

    def load_files(self):
        path1 = self.shared["file1_path"].get()
        path2 = self.shared["file2_path"].get()
        try:
            if not all([path1, path2, self.shared["start_row1"].get(), self.shared["start_row2"].get()]):
                raise ValueError("全ての項目を入力してください")

            ext1 = os.path.splitext(path1)[1].lower()
            ext2 = os.path.splitext(path2)[1].lower()

            if ext1 != ext2:
                raise ValueError("比較元と比較先のファイル形式（拡張子）が異なります")

            self.shared["file_ext"] = ext1

            if ext1 == ".csv":
                df1 = pd.read_csv(path1, header=int(self.shared["start_row1"].get()) - 1, dtype=str)
                df2 = pd.read_csv(path2, header=int(self.shared["start_row2"].get()) - 1, dtype=str)
            else:
                df1 = pd.read_excel(path1, header=int(self.shared["start_row1"].get()) - 1)
                df2 = pd.read_excel(path2, header=int(self.shared["start_row2"].get()) - 1)

            self.shared["df1"] = df1
            self.shared["df2"] = df2

            self.app.show_key_column()

        except Exception as e:
            messagebox.showerror("エラー", str(e))