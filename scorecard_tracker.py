VERSION = "1.0.1" # UPDATE THIS NUMBER ON GITHUB TO TRIGGER UPDATES
import pandas as pd
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from tksheet import Sheet # Requires pip install tksheet
import os
import glob
import json
from datetime import datetime

class ScorecardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DSP Performance Tracker Pro")
        self.root.geometry("1300x900")

        # Configuration / Settings
        self.settings_file = "settings.json"
        self.settings = {}
        self.load_settings() # Load JSON or Defaults
        
        self.df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        
        # Metrics available for advanced thresholding
        # DEFINITION: "Label": {"col": "CSV Column Name", "op": "operator", "default": value}
        # op "<": Flag if value is LESS than threshold (for Scores like FICO, Overall)
        # op ">": Flag if value is GREATER than threshold (for Rates like Speeding, Distractions)
        self.metrics = {
            "Overall Score": {"col": "Overall Score", "op": "<", "default": "90.0"},
            "FICO": {"col": "FICO Score", "op": "<", "default": "800"},
            "POD (Photo)": {"col": "POD Score", "op": "<", "default": "98.0"},
            "DCR (Completion)": {"col": "DCR Score", "op": "<", "default": "99.0"},
            "Speeding Rate": {"col": "Speeding Event Rate (per trip)", "op": ">", "default": "0.0"},
            "Seatbelt Rate": {"col": "Seatbelt-Off Rate (per trip)", "op": ">", "default": "0.0"},
            "Distractions Rate": {"col": "Distractions Rate (per trip)", "op": ">", "default": "0.0"},
            "Signal Violations": {"col": "Sign/ Signal Violations Rate (per trip)", "op": ">", "default": "0.0"},
            "Following Dist": {"col": "Following Distance Rate (per trip)", "op": ">", "default": "0.0"},
            "CDF (Customer)": {"col": "CDF DPMO Score", "op": "<", "default": "95.0"},
        }
        self.threshold_inputs = {} # Stores checkbox state
        
        self.setup_ui()
        
        # Try to load data immediately on startup
        self.load_active_data()

    def load_settings(self):
        """Load settings from JSON or create defaults based on script location."""
        # Get the directory where the script is running
        base_path = os.path.dirname(os.path.abspath(__file__))
        
        defaults = {
            "weekly_path": os.path.join(base_path, "Weekly_Scorecards"),
            "trailing_path": os.path.join(base_path, "Six_Week_Trailing")
        }

        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    loaded = json.load(f)
                    # Merge loaded with defaults in case keys are missing
                    self.settings = {**defaults, **loaded}
            except Exception as e:
                print(f"Error loading settings: {e}")
                self.settings = defaults
        else:
            self.settings = defaults
            self.save_settings() # Create the file

    def save_settings(self):
        """Save current settings to JSON file."""
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save settings: {e}")

    def setup_ui(self):
        # --- 1. SETTINGS & MODE SELECTION ---
        data_sorce_frame = ttk.LabelFrame(self.root, text="Select Data Source")
        data_sorce_frame.pack(fill="x", padx=11, pady=5)

        self.source_var = tk.StringVar(value="weekly")
        
        # Added command=self.load_active_data to reload immediately when clicked
        rb1 = ttk.Radiobutton(data_sorce_frame, text="Weekly Report", variable=self.source_var, value="weekly", command=self.load_active_data)
        rb1.grid(row=0, column=0, padx=5, pady=5)
        
        rb2 = ttk.Radiobutton(data_sorce_frame, text="6-Week Trailing", variable=self.source_var, value="trailing", command=self.load_active_data)
        rb2.grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(data_sorce_frame, text="Settings ⚙️", command=self.open_settings_menu).grid(row=0, column=2, sticky="e", padx=5, pady=5)

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
        adv_box = ttk.LabelFrame(filter_frame, text="Advanced Search (Select to Filter)")
        adv_box.pack(side="left", fill="both", expand=True, padx=5)

        # Create Grid of Checkboxes + Entry boxes
        # We assume 2 columns of controls to fit them nicely
        row_counter = 0
        col_counter = 0
        
        for label, config in self.metrics.items():
            # Variable for Checkbox
            chk_var = tk.BooleanVar()
            
            # Frame for this metric
            m_frame = ttk.Frame(adv_box)
            m_frame.grid(row=row_counter, column=col_counter, sticky="w", padx=5, pady=2)
            
            # Checkbox
            chk = ttk.Checkbutton(m_frame, text=label, variable=chk_var)
            chk.pack(side="left")
            
            # Comparison Label (e.g., "<" or ">")
            comp_text = "<" if config["op"] == "<" else ">"
            ttk.Label(m_frame, text=f" {comp_text} ").pack(side="left")
            
            # Entry for Threshold
            ent = ttk.Entry(m_frame, width=6)
            ent.insert(0, config["default"])
            ent.pack(side="left")
            
            # Store references so we can access them in apply_filters
            self.threshold_inputs[label] = (chk_var, ent)

            # Grid logic (2 items per row)
            col_counter += 1
            if col_counter > 1:
                col_counter = 0
                row_counter += 1

        btn_frame = ttk.Frame(filter_frame)
        btn_frame.pack(side="right", fill="y")
        ttk.Button(btn_frame, text="RUN QUERY", command=self.apply_filters).pack(fill="x", pady=2)
        ttk.Button(btn_frame, text="EXPORT CSV", command=self.export_data).pack(fill="x", pady=2)

        # --- 3. VIEWPORT (Using tksheet for individual cell highlighting) ---
        view_frame = ttk.Frame(self.root)
        view_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Initialize the Sheet widget
        self.sheet = Sheet(view_frame,
                           headers=[],
                           data=[[]],
                           theme="light green", # Optional theme
                           empty_horizontal=0, 
                           empty_vertical=0,
                           header_font=("Arial", 10, "bold"))
        
        # Enable basic Excel-like features
        self.sheet.enable_bindings(("single_select", 
                                    "row_select",
                                    "column_width_resize",
                                    "arrowkeys",
                                    "right_click_popup_menu",
                                    "rc_select",
                                    "copy",
                                    "cut",
                                    "paste",
                                    "delete",
                                    "undo"))
        
        self.sheet.pack(fill="both", expand=True)

        # --- 4. SUMMARY ---
        self.summary_frame = ttk.LabelFrame(self.root, text="Totals & Averages")
        self.summary_frame.pack(fill="x", padx=10, pady=10)
        self.summary_label = ttk.Label(self.summary_frame, text="Load folders and run query.", font=('Arial', 10, 'bold'))
        self.summary_label.pack(pady=10)

    def open_settings_menu(self):
        """Opens a popup window to manage folder paths."""
        top = tk.Toplevel(self.root)
        top.title("Folder Settings")
        top.geometry("600x200")

        # Weekly Folder
        ttk.Label(top, text="Weekly Scorecards Folder:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        entry_weekly = ttk.Entry(top, width=50)
        entry_weekly.insert(0, self.settings["weekly_path"])
        entry_weekly.grid(row=0, column=1, padx=10)
        
        def browse_weekly():
            path = filedialog.askdirectory()
            if path:
                entry_weekly.delete(0, tk.END)
                entry_weekly.insert(0, path)
        
        ttk.Button(top, text="Browse", command=browse_weekly).grid(row=0, column=2, padx=10)

        # Trailing Folder
        ttk.Label(top, text="6-Week Trailing Folder:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        entry_trailing = ttk.Entry(top, width=50)
        entry_trailing.insert(0, self.settings["trailing_path"])
        entry_trailing.grid(row=1, column=1, padx=10)

        def browse_trailing():
            path = filedialog.askdirectory()
            if path:
                entry_trailing.delete(0, tk.END)
                entry_trailing.insert(0, path)

        ttk.Button(top, text="Browse", command=browse_trailing).grid(row=1, column=2, padx=10)

        # Save Button
        def save_and_close():
            self.settings["weekly_path"] = entry_weekly.get()
            self.settings["trailing_path"] = entry_trailing.get()
            self.save_settings()
            self.load_active_data() # Reload data with new paths
            top.destroy()
            messagebox.showinfo("Settings", "Settings saved and data reloaded.")

        ttk.Button(top, text="Save Settings", command=save_and_close).grid(row=2, column=1, pady=20)

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
        path = self.settings.get(f"{mode}_path", "")
        
        # Visual feedback on data source frame title or summary
        if not path or not os.path.exists(path):
            self.summary_label.config(text=f"Error: Folder not found for '{mode}'. Check Settings.")
            self.df = pd.DataFrame() # Clear data
            self.update_viewport()
            return

        files = glob.glob(os.path.join(path, "*.csv"))
        all_dfs = []
        for f in files:
            try:
                temp_df = pd.read_csv(f, encoding='utf-8-sig')
                temp_df.columns = temp_df.columns.str.strip()
                all_dfs.append(temp_df)
            except Exception as e:
                print(f"Error reading {f}: {e}")

        if all_dfs:
            self.df = pd.concat(all_dfs, ignore_index=True)
            
            # --- FIX IS HERE ---
            # Pre-clean numeric columns for filtering
            # We iterate through the metric configurations and extract the "col" string
            for config in self.metrics.values():
                col_name = config["col"]
                if col_name in self.df.columns:
                    self.df[col_name] = self.df[col_name].apply(self.clean_numeric)
            # -------------------
            
            # Update DA list
            if "Delivery Associate" in self.df.columns:
                employees = sorted(self.df["Delivery Associate"].astype(str).unique().tolist())
                self.employee_dropdown['values'] = ["ALL"] + employees
                self.employee_dropdown.set("ALL")
            
            self.summary_label.config(text=f"Loaded {len(self.df)} rows from {mode} folder. Ready to query.")
        else:
            self.df = pd.DataFrame()
            self.summary_label.config(text=f"No CSV files found in {path}")
        
        # Clear previous results when switching sources
        self.filtered_df = pd.DataFrame()
        self.update_viewport()

    def apply_filters(self):
        if self.df.empty:
            self.load_active_data()
            if self.df.empty: return

        query_df = self.df.copy()

        # 1. Date to Week conversion
        if "Week" in query_df.columns:
            try:
                start_w = self.date_to_iso_week(self.cal_start.get_date())
                end_w = self.date_to_iso_week(self.cal_end.get_date())
                query_df = query_df[(query_df["Week"] >= start_w) & (query_df["Week"] <= end_w)]
            except: pass

        # 2. Employee
        emp = self.employee_var.get()
        if emp != "ALL" and "Delivery Associate" in query_df.columns:
            query_df = query_df[query_df["Delivery Associate"] == emp]

        # 3. Advanced Thresholds
        # NOTE: We do NOT remove rows here anymore based on threshold.
        # We pass the Date/Employee filtered data to viewport, 
        # and let viewport handle the logic of hiding rows that don't match criteria.
        
        self.filtered_df = query_df
        self.update_viewport()
        self.calculate_summary()

    def update_viewport(self):
        # Clear existing data and highlights
        self.sheet.set_sheet_data([[]])
        self.sheet.dehighlight_all()
        
        if self.filtered_df.empty: return

        # --- 1. DETERMINE COLUMNS ---
        all_cols = list(self.filtered_df.columns)
        
        # Exclude keywords (cleaning up the view)
        exclude_keywords = ['Tier', 'Platinum', 'Gold', 'Bronze', 'Silver', 'Standing', 'Transporter ID', 'Weight Applied', "PSB", "PSB Score"]
        view_cols = [c for c in all_cols if not any(k in c for k in exclude_keywords)]
        
        # Ensure key identifiers exist
        if "Week" not in view_cols and "Week" in all_cols: view_cols.insert(0, "Week")
        if "Delivery Associate" not in view_cols and "Delivery Associate" in all_cols: view_cols.insert(1, "Delivery Associate")

        # --- 2. PREPARE DATA & FILTERING ---
        final_rows = []
        
        # We need to check filtering logic: 
        # If checkboxes are active, we ONLY show rows that have at least one violation.
        filters_active = any(chk.get() for chk, _ in self.threshold_inputs.values())

        # Map metric column names to their settings for fast lookup
        # Format: {'Overall Score': {'op': '<', 'limit': 90.0}, ...}
        active_thresholds = {}
        for label, (chk_var, ent_widget) in self.threshold_inputs.items():
            # We track thresholds even if checkbox is NOT checked, so we can highlight cells
            # But we use the checkbox state later to decide if we hide the row
            col_name = self.metrics[label]["col"]
            operator = self.metrics[label]["op"]
            try:
                limit = float(ent_widget.get())
            except:
                limit = 0.0
            
            active_thresholds[col_name] = {
                "op": operator, 
                "limit": limit, 
                "checked": chk_var.get()
            }

        # Iterate Dataframe
        for _, row in self.filtered_df.iterrows():
            has_violation_for_filter = False
            
            # Check if this row should be shown (based on checked boxes)
            if filters_active:
                for col_name, config in active_thresholds.items():
                    if config["checked"] and col_name in row:
                        val = row[col_name]
                        if config["op"] == "<" and val < config["limit"]:
                            has_violation_for_filter = True
                            break
                        elif config["op"] == ">" and val > config["limit"]:
                            has_violation_for_filter = True
                            break
                
                if not has_violation_for_filter:
                    continue # Skip this row

            # Prepare row data for display (only visible columns)
            row_data = [row[c] for c in view_cols]
            final_rows.append(row_data)

        # --- 3. POPULATE SHEET ---
        self.sheet.headers(view_cols)
        self.sheet.set_sheet_data(final_rows)
        
        # --- 4. HIGHLIGHT INDIVIDUAL CELLS ---
        # Loop through the data we just displayed
        bg_color = "#ffcccc" # Light Red
        
        for r_idx, row_list in enumerate(final_rows):
            for c_idx, cell_value in enumerate(row_list):
                col_name = view_cols[c_idx]
                
                # Check if this column has a threshold rule
                if col_name in active_thresholds:
                    config = active_thresholds[col_name]
                    
                    # Ensure value is numeric for comparison
                    try:
                        val = float(cell_value)
                        
                        # Apply Highlight Logic
                        # Note: We highlight if it violates the number, 
                        # regardless of whether the specific checkbox was checked for filtering rows.
                        if config["op"] == "<" and val < config["limit"]:
                             self.sheet.highlight_cells(row=r_idx, column=c_idx, bg=bg_color)
                        elif config["op"] == ">" and val > config["limit"]:
                             self.sheet.highlight_cells(row=r_idx, column=c_idx, bg=bg_color)
                    except:
                        pass # Non-numeric data in a metric column? Skip.

        # Update Summary
        self.summary_label.config(text=f"Rows Displayed: {len(final_rows)}")
        
        # Auto-fit columns (tksheet built-in)
        self.sheet.set_all_cell_sizes_to_text_width()
        
    def calculate_summary(self):
        if self.filtered_df.empty:
            self.summary_label.config(text="No results.")
            return

        txt_parts = [f"Filtered Count: {len(self.filtered_df)}"]

        if "Packages Delivered" in self.filtered_df.columns:
            pkgs = self.filtered_df["Packages Delivered"].apply(self.clean_numeric).sum()
            txt_parts.append(f"Total Packages: {int(pkgs):,}")
        
        if "Overall Score" in self.filtered_df.columns:
            avg_score = self.filtered_df["Overall Score"].mean()
            txt_parts.append(f"Avg Score: {avg_score:.2f}")

        if "POD Score" in self.filtered_df.columns:
            avg_pod = self.filtered_df["POD Score"].mean()
            txt_parts.append(f"Avg POD: {avg_pod:.2f}")

        if "FICO Score" in self.filtered_df.columns:
            # Handle zeros in FICO if needed, usually we don't count 0
            fico_clean = self.filtered_df["FICO Score"].replace(0, pd.NA).dropna()
            if not fico_clean.empty:
                avg_fico = fico_clean.mean()
                txt_parts.append(f"Avg FICO: {avg_fico:.1f}")

        self.summary_label.config(text=" | ".join(txt_parts))

    def export_data(self):
        if self.filtered_df.empty: return
        fp = filedialog.asksaveasfilename(defaultextension=".csv")
        if fp:
            self.filtered_df.to_csv(fp, index=False)
            messagebox.showinfo("Success", "Exported!")

def main():
    root = tk.Tk()
    app = ScorecardApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()