import tkinter as tk
from gui.file_select_frame import FileSelectFrame
from gui.key_column_frame import KeyColumnFrame
from gui.column_mapping_frame import ColumnMappingFrame

class Setting:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV差分比較ツール")

        self.shared_data = {
            "file1_path": tk.StringVar(),
            "file2_path": tk.StringVar(),
            "start_row1": tk.StringVar(),
            "start_row2": tk.StringVar(),
            "key_col1": tk.StringVar(),
            "key_col2": tk.StringVar(),
            "col_mapping": {},
            "file_ext": None,
            "df1": None,
            "df2": None,
        }

        self.show_file_select()

    def clear_root(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def show_file_select(self):
        self.clear_root()
        FileSelectFrame(self)

    def show_key_column(self):
        self.clear_root()
        KeyColumnFrame(self)

    def show_column_mapping(self):
        self.clear_root()
        ColumnMappingFrame(self)