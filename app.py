import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import polars as pl
import os
from ttkbootstrap import Style

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Filter Tool")
        self.root.geometry("1000x700")
        self.style = Style("cyborg")
        self.file_path = None
        self.df = None

        self.filter_widgets = {}

        self.create_widgets()

    def create_widgets(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        load_btn = ttk.Button(top_frame, text="Load Excel File", command=self.load_file)
        load_btn.pack(side=tk.LEFT)

        export_btn = ttk.Button(top_frame, text="Export Filtered Data", command=self.export_data)
        export_btn.pack(side=tk.LEFT, padx=10)

        self.filter_frame = ttk.Labelframe(self.root, text="Filters")
        self.filter_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=10)

        apply_btn = ttk.Button(self.root, text="Apply Filters", command=self.apply_filters)
        apply_btn.pack(pady=10)

        self.result_label = ttk.Label(self.root, text="")
        self.result_label.pack()

        self.table_frame = ttk.Frame(self.root)
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = None

    def load_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file:
            return

        try:
            self.df = pl.read_excel(file)
            self.file_path = file
            self.setup_filters()
            self.result_label.config(text="File loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def setup_filters(self):
        for widget in self.filter_frame.winfo_children():
            widget.destroy()

        self.filter_widgets.clear()

        for i, col in enumerate(self.df.columns):
            label = ttk.Label(self.filter_frame, text=col)
            label.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)

            entry = ttk.Entry(self.filter_frame)
            entry.grid(row=i, column=1, padx=5, pady=2)
            self.filter_widgets[col] = entry

    def apply_filters(self):
        if self.df is None:
            return

        filtered_df = self.df

        for col, widget in self.filter_widgets.items():
            val = widget.get().strip()
            if val:
                filtered_df = filtered_df.filter(pl.col(col).cast(str).str.contains(val, case=False))

        self.filtered_df = filtered_df

        self.show_table(filtered_df)
        self.result_label.config(text=f"Filtered Rows: {len(filtered_df)}")

    def show_table(self, df):
        if self.tree:
            self.tree.destroy()

        self.tree = ttk.Treeview(self.table_frame, columns=df.columns, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")

        for row in df.iter_rows():
            self.tree.insert("", tk.END, values=row)

    def export_data(self):
        if not hasattr(self, 'filtered_df'):
            messagebox.showwarning("Warning", "No filtered data to export!")
            return

        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file:
            return

        try:
            filtered_pd = self.filtered_df.to_pandas()
            filtered_pd.to_excel(file, index=False)
            messagebox.showinfo("Success", "Data exported successfully!")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()
