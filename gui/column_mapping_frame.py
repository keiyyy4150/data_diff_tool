import tkinter as tk
from tkinter import ttk, messagebox
from logic.exporter import export_diff

class ColumnMappingFrame(tk.Frame):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.shared = app.shared_data
        self.combos = {}
        self.pack(padx=10, pady=10)
        self.create_widgets()

    def create_widgets(self):
        df1 = self.shared["df1"]
        df2 = self.shared["df2"]

        tk.Label(self, text="比較元カラム", width=30).grid(row=0, column=0, sticky="w")
        tk.Label(self, text="比較先カラム", width=30).grid(row=0, column=1, sticky="w")

        for i, col1 in enumerate(df1.columns):
            tk.Label(self, text=col1).grid(row=i+1, column=0, sticky="w")
            combo = ttk.Combobox(self, values=list(df2.columns), state="readonly")
            combo.grid(row=i+1, column=1, sticky="w")
            if col1 in df2.columns:
                combo.set(col1)
            self.combos[col1] = combo

        button_row = i + 2
        tk.Button(self, text="戻る", command=self.app.show_key_column).grid(row=button_row, column=0, pady=10, sticky="w")
        tk.Button(self, text="終了", command=self.app.root.quit).grid(row=button_row, column=1, pady=10, sticky="e")
        tk.Button(self, text="実行する", command=self.execute).grid(row=button_row, column=2, pady=10, sticky="e")

    def execute(self):
        for k, v in self.combos.items():
            self.shared["col_mapping"][k] = v.get()

        try:
            export_diff(self.shared)
            messagebox.showinfo("完了", "差分ファイルを出力しました。")
            self.app.root.destroy()
        except Exception as e:
            messagebox.showerror("エラー", str(e))