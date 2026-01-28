import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import os
import glob
from datetime import datetime

class ScorecardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DSP Performance Tracker Pro")
        self.root.geometry("1300x900")

        # Configuration / Settings
        self.settings = {"weekly_path": "", "trailing_path": ""}
        self.df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        
        # Metrics available for advanced thresholding
        self.metrics = {
            "Overall Score": "Overall Score",
            "FICO": "FICO Score",
            "Speeding": "Speeding Event Rate Score",
            "Seatbelt": "Seatbelt-Off Rate Score",
            "Distractions": "Distractions Rate Score",
            "Signal Violations": "Sign/ Signal Violations Rate Score",
            "Following Distance": "Following Distance Rate Score",
            "CDF (Customer Delivery)": "CDF DPMO Score",
            "DCR (Delivery Completion)": "DCR Score",
            "POD (Photo on Delivery)": "POD Score"
        }
        self.threshold_vars = {} # Stores checkbox state
        
        self.setup_ui()

    def setup_ui(self):
        # --- 1. SETTINGS & MODE SELECTION ---
        data_sorce_frame = ttk.LabelFrame(self.root, text="Select Data Source")
        data_sorce_frame.pack(fill="x", padx=11, pady=5)

        # ttk.Button(settings_frame, text="Set Weekly Folder", command=lambda: self.set_folder("weekly")).grid(row=0, column=0, padx=5, pady=5)
        # self.lbl_weekly = ttk.Label(settings_frame, text="Not set", foreground="gray")
        # self.lbl_weekly.grid(row=0, column=1, padx=5)

        # ttk.Button(settings_frame, text="Set Trailing Folder", command=lambda: self.set_folder("trailing")).grid(row=0, column=2, padx=5, pady=5)
        # self.lbl_trailing = ttk.Label(settings_frame, text="Not set", foreground="gray")
        # self.lbl_trailing.grid(row=0, column=3, padx=5)

        self.source_var = tk.StringVar(value="weekly")
        ttk.Radiobutton(data_sorce_frame, text="Weekly Report", variable=self.source_var, value="weekly").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(data_sorce_frame, text="6-Week Trailing", variable=self.source_var, value="trailing").grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(data_sorce_frame, text="Settings ⚙️", command=lambda: self.set_folder("trailing")).grid(row=0, column=2, sticky="e", padx=5, pady=5)

        # --- 2. MAIN FILTERS ---
        filter_frame = ttk.Frame(self.root)
        filter_frame.pack(fill="x", padx=10, pady=5)

        # Basic Filter Box
        basic_box = ttk.LabelFrame(filter_frame, text="General Filters")
        basic_box.pack(side="left", fill="y", padx=5)

        ttk.Label(basic_box, text="Employee:").grid(row=0, column=0, padx=5, pady=5)
        self.employee_var = tk.StringVar()
        self.employee_dropdown = ttk.Combobox(basic_box, textvariable=self.employee_var, width=25)
        self.employee_dropdown.grid(row=0, column=1, padx=5)

        ttk.Label(basic_box, text="Start Date:").grid(row=1, column=0, padx=5)
        self.cal_start = DateEntry(basic_box, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.cal_start.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(basic_box, text="End Date:").grid(row=2, column=0, padx=5)
        self.cal_end = DateEntry(basic_box, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.cal_end.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        # Advanced Threshold Box
        adv_box = ttk.LabelFrame(filter_frame, text="Advanced Search (Scores BELOW Threshold)")
        adv_box.pack(side="left", fill="both", expand=True, padx=5)

        ttk.Label(adv_box, text="Threshold Value:").grid(row=0, column=0, padx=5)
        self.threshold_val = ttk.Entry(adv_box, width=8)
        self.threshold_val.insert(0, "90.0")
        self.threshold_val.grid(row=0, column=1, sticky="w")

        # Grid of checkboxes for metrics
        for i, (label, col_name) in enumerate(self.metrics.items()):
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(adv_box, text=label, variable=var)
            chk.grid(row=(i // 5) + 1, column=i % 5, sticky="w", padx=10)
            self.threshold_vars[col_name] = var

        btn_frame = ttk.Frame(filter_frame)
        btn_frame.pack(side="right", fill="y")
        ttk.Button(btn_frame, text="RUN QUERY", command=self.apply_filters).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="EXPORT CSV", command=self.export_data).pack(fill="x", pady=2)

        # --- 3. VIEWPORT ---
        view_frame = ttk.Frame(self.root)
        view_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(view_frame, show="headings")
        sy = ttk.Scrollbar(view_frame, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(view_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.pack(side="top", fill="both", expand=True)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")

        # --- 4. SUMMARY ---
        self.summary_frame = ttk.LabelFrame(self.root, text="Totals & Averages")
        self.summary_frame.pack(fill="x", padx=10, pady=10)
        self.summary_label = ttk.Label(self.summary_frame, text="Load folders and run query.", font=('Arial', 10, 'bold'))
        self.summary_label.pack(pady=10)

    def set_folder(self, mode):
        path = filedialog.askdirectory()
        if path:
            if mode == "weekly":
                self.settings["weekly_path"] = path
                self.lbl_weekly.config(text=path, foreground="black")
            else:
                self.settings["trailing_path"] = path
                self.lbl_trailing.config(text=path, foreground="black")
            self.load_active_data()

    def date_to_iso_week(self, dt):
        """Converts a datetime object to 2026-W01 format."""
        year, week, day = dt.isocalendar()
        return f"{year}-W{week:02d}"

    def clean_numeric(self, val):
        if pd.isna(val) or val == "": return 0.0
        if isinstance(val, str):
            val = val.replace('%', '').replace(',', '')
        try:
            return float(val)
        except:
            return 0.0

    def load_active_data(self):
        mode = self.source_var.get()
        path = self.settings[f"{mode}_path"]
        if not path: return

        files = glob.glob(os.path.join(path, "*.csv"))
        all_dfs = []
        for f in files:
            try:
                temp_df = pd.read_csv(f, encoding='utf-8-sig')
                temp_df.columns = temp_df.columns.str.strip()
                all_dfs.append(temp_df)
            except Exception as e:
                print(f"Error: {e}")

        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            # Pre-clean numeric columns for filtering
            for col in self.metrics.values():
                if col in self.df.columns:
                    self.df[col] = self.df[col].apply(self.clean_numeric)
            
            # Update DA list
            employees = sorted(self.df["Delivery Associate"].unique().tolist())
            self.employee_dropdown['values'] = ["ALL"] + employees
            self.employee_dropdown.set("ALL")

    def apply_filters(self):
        # Always reload data from the selected folder to ensure fresh results
        self.load_active_data()
        
        if self.df.empty:
            messagebox.showerror("Error", "No data loaded for this mode.")
            return

        query_df = self.df.copy()

        # 1. Date to Week conversion
        start_w = self.date_to_iso_week(self.cal_start.get_date())
        end_w = self.date_to_iso_week(self.cal_end.get_date())
        query_df = query_df[(query_df["Week"] >= start_w) & (query_df["Week"] <= end_w)]

        # 2. Employee
        emp = self.employee_var.get()
        if emp != "ALL":
            query_df = query_df[query_df["Delivery Associate"] == emp]

        # 3. Advanced Thresholds (Below X)
        try:
            limit = float(self.threshold_val.get())
        except:
            limit = 100.0

        for col, var in self.threshold_vars.items():
            if var.get(): # If checkbox is checked
                if col in query_df.columns:
                    query_df = query_df[query_df[col] < limit]

        self.filtered_df = query_df
        self.update_viewport()
        self.calculate_summary()

    def update_viewport(self):
        self.tree.delete(*self.tree.get_children())
        if self.filtered_df.empty: return

        cols = list(self.filtered_df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in self.filtered_df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def calculate_summary(self):
        if self.filtered_df.empty:
            self.summary_label.config(text="No results.")
            return

        # Sums
        pkgs = self.filtered_df["Packages Delivered"].apply(self.clean_numeric).sum()
        
        # Averages
        avg_score = self.filtered_df["Overall Score"].mean()
        avg_pod = self.filtered_df["POD Score"].mean()
        avg_fico = self.filtered_df["FICO Score"].replace(0, pd.NA).dropna().mean()

        txt = (f"Filtered Count: {len(self.filtered_df)} | "
            f"Total Packages: {int(pkgs):,} | "
            f"Avg Score: {avg_score:.2f} | "
            f"Avg POD Score: {avg_pod:.2f} | "
            f"Avg FICO: {avg_fico:.1f}")
        self.summary_label.config(text=txt)

    def export_data(self):
        if self.filtered_df.empty: return
        fp = filedialog.asksaveasfilename(defaultextension=".csv")
        if fp:
            self.filtered_df.to_csv(fp, index=False)
            messagebox.showinfo("Success", "Exported!")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScorecardApp(root)
    root.mainloop()