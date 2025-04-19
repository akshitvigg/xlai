import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
import os

class ExcelSearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Search Tool")
        self.root.geometry("800x600")
        self.root.configure(padx=20, pady=20)
        
        self.dataframe = None
        self.filtered_df = None
        self.filters = {}
        self.filter_widgets = {}
        
        self.create_widgets()
        
    def create_widgets(self):
        file_frame = ttk.LabelFrame(self.root, text="Select Excel File")
        file_frame.pack(fill="x", padx=10, pady=10)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).pack(side=tk.LEFT, padx=5, pady=10, expand=True, fill="x")
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT, padx=5, pady=10)
        ttk.Button(file_frame, text="Load File", command=self.load_file).pack(side=tk.LEFT, padx=5, pady=10)
        
        self.filter_frame = ttk.LabelFrame(self.root, text="Search Filters")
        self.filter_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.results_var = tk.StringVar(value="No file loaded")
        ttk.Label(self.root, textvariable=self.results_var, font=('Arial', 10, 'bold')).pack(pady=5)
        
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(button_frame, text="Apply Filters", command=self.apply_filters).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset Filters", command=self.reset_filters).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export Results", command=self.export_results).pack(side=tk.LEFT, padx=5)
        
        results_frame = ttk.LabelFrame(self.root, text="Results Preview (First 100 rows)")
        results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tree = ttk.Treeview(results_frame)
        self.tree.pack(fill="both", expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
    def browse_file(self):
        filepath = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepath:
            self.file_path_var.set(filepath)
    
    def load_file(self):
        filepath = self.file_path_var.get()
        if not filepath:
            return
            
        try:
            self.dataframe = pd.read_excel(filepath)
            self.filtered_df = self.dataframe.copy()
            
            self.update_results_count()
            
            for widget in self.filter_frame.winfo_children():
                widget.destroy()
            
            self.filters = {}
            self.filter_widgets = {}
            
            self.create_filter_widgets()
            
            self.update_treeview()
            
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to load file: {str(e)}")
    
    def create_filter_widgets(self):
        canvas = tk.Canvas(self.filter_frame)
        scrollbar = ttk.Scrollbar(self.filter_frame, orient="vertical", command=canvas.yview)
        
        filter_container = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        
        canvas.create_window((0, 0), window=filter_container, anchor="nw")
        filter_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        row = 0
        for col in self.dataframe.columns:
            unique_vals = self.dataframe[col].dropna().unique()
            
            filter_frame = ttk.Frame(filter_container)
            filter_frame.grid(row=row, column=0, sticky="w", padx=5, pady=5)
            
            ttk.Label(filter_frame, text=f"{col}:").grid(row=0, column=0, sticky="w", padx=5)
            
            if len(unique_vals) < 30 and pd.api.types.is_object_dtype(self.dataframe[col]):
                cb_frame = ttk.Frame(filter_frame)
                cb_frame.grid(row=0, column=1, sticky="w")
                
                select_all_var = tk.BooleanVar(value=False)
                select_all_cb = ttk.Checkbutton(
                    cb_frame, 
                    text="Select All / None", 
                    variable=select_all_var,
                    command=lambda c=col, v=select_all_var: self.toggle_all(c, v)
                )
                select_all_cb.pack(anchor="w")
                
                self.filter_widgets[col] = {
                    "type": "checkbox",
                    "select_all": select_all_var,
                    "values": {}
                }
                
                for val in sorted(unique_vals):
                    if pd.notna(val):
                        var = tk.BooleanVar(value=False)
                        cb = ttk.Checkbutton(
                            cb_frame, 
                            text=str(val), 
                            variable=var,
                            command=lambda c=col: self.update_select_all(c)
                        )
                        cb.pack(anchor="w")
                        self.filter_widgets[col]["values"][str(val)] = var
            
            else:
                var = tk.StringVar()
                entry = ttk.Entry(filter_frame, textvariable=var, width=30)
                entry.grid(row=0, column=1, sticky="w", padx=5)
                
                self.filter_widgets[col] = {
                    "type": "text",
                    "widget": entry,
                    "var": var
                }
                
                ttk.Label(
                    filter_frame, 
                    text="Enter search terms separated by commas", 
                    font=("Arial", 8, "italic")
                ).grid(row=1, column=1, sticky="w", padx=5)
                
            row += 1
    
    def toggle_all(self, column, select_all_var):
        is_selected = select_all_var.get()
        for var in self.filter_widgets[column]["values"].values():
            var.set(is_selected)
    
    def update_select_all(self, column):
        all_selected = all(var.get() for var in self.filter_widgets[column]["values"].values())
        none_selected = not any(var.get() for var in self.filter_widgets[column]["values"].values())
        
        self.filter_widgets[column]["select_all"].set(all_selected)
    
    def apply_filters(self):
        self.filtered_df = self.dataframe.copy()
        
        for col, widget_info in self.filter_widgets.items():
            if widget_info["type"] == "checkbox":
                selected_values = [
                    val for val, var in widget_info["values"].items() 
                    if var.get()
                ]
                
                if selected_values:
                    self.filtered_df = self.filtered_df[self.filtered_df[col].astype(str).isin(selected_values)]
            
            elif widget_info["type"] == "text":
                search_text = widget_info["var"].get().strip()
                if search_text:
                    search_terms = [term.strip() for term in search_text.split(",")]
                    
                    mask = self.filtered_df[col].astype(str).str.contains(
                        '|'.join(search_terms), 
                        case=False, 
                        na=False
                    )
                    self.filtered_df = self.filtered_df[mask]
        
        self.update_results_count()
        self.update_treeview()
    
    def reset_filters(self):
        for col, widget_info in self.filter_widgets.items():
            if widget_info["type"] == "checkbox":
                widget_info["select_all"].set(False)
                for var in widget_info["values"].values():
                    var.set(False)
            elif widget_info["type"] == "text":
                widget_info["var"].set("")
        
        self.filtered_df = self.dataframe.copy()
        
        self.update_results_count()
        self.update_treeview()
    
    def update_results_count(self):
        if self.filtered_df is not None:
            self.results_var.set(f"Showing {len(self.filtered_df)} of {len(self.dataframe)} total rows")
    
    def update_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.tree["columns"] = list(self.filtered_df.columns)
        
        self.tree.column("#0", width=0, stretch=tk.NO)
        for col in self.filtered_df.columns:
            self.tree.column(col, anchor=tk.W, width=100)
            self.tree.heading(col, text=col, anchor=tk.W)
        
        preview_df = self.filtered_df.head(100)
        for i, row in preview_df.iterrows():
            values = [row[col] for col in self.filtered_df.columns]
            self.tree.insert("", tk.END, text="", values=values)
    
    def export_results(self):
        if self.filtered_df is None or len(self.filtered_df) == 0:
            tk.messagebox.showinfo("Export", "No data to export.")
            return
            
        filepath = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*"))
        )
        
        if not filepath:
            return
            
        try:
            _, ext = os.path.splitext(filepath)
            
            if ext.lower() == '.csv':
                self.filtered_df.to_csv(filepath, index=False)
            else:
                self.filtered_df.to_excel(filepath, index=False)
                
            tk.messagebox.showinfo("Export", f"Data successfully exported to {filepath}")
            
        except Exception as e:
            tk.messagebox.showerror("Export Error", f"Failed to export data: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSearchTool(root)
    root.mainloop()