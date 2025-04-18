import tkinter as tk
from tkinter import filedialog, ttk
from ttkbootstrap import Style
import polars as pl
from tkinter import messagebox

class ExcelGodMode:
    def __init__(self, root):
        self.root = root
        self.root.title("📂 Excel God Mode Viewer")
        self.style = Style(theme="darkly")
        self.root.geometry("1000x600")

        self.df = None
        self.filter_vars = {}

        # Top bar
        top_frame = ttk.Frame(root)
        top_frame.pack(fill='x', pady=10)

        ttk.Button(top_frame, text="🧾 Upload Excel", command=self.load_file).pack(side='left', padx=5)
        ttk.Button(top_frame, text="🧹 Clear Filters", command=self.clear_filters).pack(side='left', padx=5)
        ttk.Button(top_frame, text="💾 Download Filtered", command=self.download_filtered).pack(side='left', padx=5)

        # Filter frame
        self.frame_filters = ttk.Frame(root)
        self.frame_filters.pack(fill='x', pady=5)

        # Treeview frame with scrollbars
        tree_frame = ttk.Frame(root)
        tree_frame.pack(expand=True, fill='both', padx=10, pady=10)

        self.tree = ttk.Treeview(tree_frame, show='headings')
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            self.df = pl.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Read Error", f"Failed to load Excel: {e}")
            return

        self.build_filter_fields()
        self.update_treeview(self.df)

    def build_filter_fields(self):
        for widget in self.frame_filters.winfo_children():
            widget.destroy()

        self.filter_vars.clear()

        for col in self.df.columns:
            frame = ttk.Frame(self.frame_filters)
            frame.pack(side='left', padx=4)
            ttk.Label(frame, text=col).pack()
            var = tk.StringVar()
            entry = ttk.Entry(frame, textvariable=var, width=15)
            entry.pack()
            entry.bind("<KeyRelease>", lambda e: self.apply_filters())
            self.filter_vars[col] = var

    def apply_filters(self):
        if self.df is None:
            return

        filtered_df = self.df

        for col, var in self.filter_vars.items():
            val = var.get().strip().lower()
            if val:
                try:
                    filtered_df = filtered_df.filter(
                        pl.col(col).cast(str).str.to_lowercase().str.contains(val)
                    )
                except Exception as e:
                    print(f"Filter error in {col}: {e}")

        self.update_treeview(filtered_df)

    def update_treeview(self, data):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = data.columns

        for col in data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor='w')

        for row in data.iter_rows(named=True):
            self.tree.insert("", "end", values=list(row.values()))

        self.filtered_df = data

    def clear_filters(self):
        for var in self.filter_vars.values():
            var.set("")
        self.update_treeview(self.df)

    def download_filtered(self):
        if not hasattr(self, 'filtered_df') or self.filtered_df is None:
            messagebox.showwarning("No Data", "No filtered data to save.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 title="Save filtered data as")
        if save_path:
            try:
                self.filtered_df.write_excel(save_path)
                messagebox.showinfo("Saved", f"Filtered data saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelGodMode(root)
    root.mainloop()
