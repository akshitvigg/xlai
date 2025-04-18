import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Filter Tool")
        self.df = None
        self.filter_entries = {}

        self.setup_ui()

    def setup_ui(self):
        # Load button
        load_btn = tk.Button(self.root, text="Upload Excel File", command=self.load_file)
        load_btn.pack(pady=10)

        # Frame for filters
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack(pady=10)

        # Search button
        self.search_btn = tk.Button(self.root, text="Search", command=self.apply_filters, state='disabled')
        self.search_btn.pack(pady=5)

        # Table
        self.tree = ttk.Treeview(self.root)
        self.tree.pack(padx=10, pady=10, fill='both', expand=True)

        # Export button
        self.export_btn = tk.Button(self.root, text="Export Filtered Data", command=self.export_data, state='disabled')
        self.export_btn.pack(pady=5)

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        try:
            self.df = pd.read_excel(filepath)
            self.build_filters()
            self.populate_table(self.df)
            self.search_btn.config(state='normal')
            self.export_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def build_filters(self):
        # Clear previous filters
        for widget in self.filter_frame.winfo_children():
            widget.destroy()
        self.filter_entries.clear()

        for idx, column in enumerate(self.df.columns):
            tk.Label(self.filter_frame, text=column).grid(row=idx, column=0, padx=5, pady=2, sticky='w')
            entry = tk.Entry(self.filter_frame)
            entry.grid(row=idx, column=1, padx=5, pady=2)
            self.filter_entries[column] = entry

    def apply_filters(self):
        filtered_df = self.df.copy()
        for column, entry in self.filter_entries.items():
            val = entry.get().strip()
            if val:
                filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(val, case=False, na=False)]
        self.populate_table(filtered_df)

    def populate_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def export_data(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            data = [self.tree.item(item)["values"] for item in self.tree.get_children()]
            filtered_df = pd.DataFrame(data, columns=self.df.columns)
            filtered_df.to_excel(file, index=False)
            messagebox.showinfo("Success", "Filtered data exported successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1000x600")
    app = ExcelFilterApp(root)
    root.mainloop()
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Filter Tool")
        self.df = None
        self.filter_entries = {}

        self.setup_ui()

    def setup_ui(self):

        load_btn = tk.Button(self.root, text="Upload Excel File", command=self.load_file)
        load_btn.pack(pady=10)

      
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack(pady=10)

        
        self.search_btn = tk.Button(self.root, text="Search", command=self.apply_filters, state='disabled')
        self.search_btn.pack(pady=5)

    
        self.tree = ttk.Treeview(self.root)
        self.tree.pack(padx=10, pady=10, fill='both', expand=True)

      
        self.export_btn = tk.Button(self.root, text="Export Filtered Data", command=self.export_data, state='disabled')
        self.export_btn.pack(pady=5)

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        try:
            self.df = pd.read_excel(filepath)
            self.build_filters()
            self.populate_table(self.df)
            self.search_btn.config(state='normal')
            self.export_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def build_filters(self):

        for widget in self.filter_frame.winfo_children():
            widget.destroy()
        self.filter_entries.clear()

        for idx, column in enumerate(self.df.columns):
            tk.Label(self.filter_frame, text=column).grid(row=idx, column=0, padx=5, pady=2, sticky='w')
            entry = tk.Entry(self.filter_frame)
            entry.grid(row=idx, column=1, padx=5, pady=2)
            self.filter_entries[column] = entry

    def apply_filters(self):
        filtered_df = self.df.copy()
        for column, entry in self.filter_entries.items():
            val = entry.get().strip()
            if val:
                filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(val, case=False, na=False)]
        self.populate_table(filtered_df)

    def populate_table(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def export_data(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            data = [self.tree.item(item)["values"] for item in self.tree.get_children()]
            filtered_df = pd.DataFrame(data, columns=self.df.columns)
            filtered_df.to_excel(file, index=False)
            messagebox.showinfo("Success", "Filtered data exported successfully!")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1000x600")
    app = ExcelFilterApp(root)
    root.mainloop()
