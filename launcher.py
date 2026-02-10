import os
import sys
import requests
import importlib.util
import threading
import re

# We explicitly import these so PyInstaller bundles them for the external script
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkFont 
import pandas
import numpy
import openpyxl
import tkcalendar
import tksheet
import babel.numbers

# --- CONFIGURATION ---
# IMPORTANT: Use the "Raw" view URL from GitHub
GITHUB_RAW_URL = "https://raw.githubusercontent.com/benberryhill/Scorecard-Tracker/refs/heads/working_branch/scorecard_tracker.py"
SCRIPT_NAME = "scorecard_tracker.py"

class Launcher:
    def __init__(self):
        self.ensure_folders_exist()
        self.root = tk.Tk()
        self.root.overrideredirect(True) # No title bar (Splash screen style)
        self.root.geometry("300x100+500+300")
        
        # Center on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width/2) - (300/2))
        y = int((screen_height/2) - (100/2))
        self.root.geometry(f"300x100+{x}+{y}")

        # UI
        frame = ttk.Frame(self.root, relief="raised", borderwidth=2)
        frame.pack(fill="both", expand=True)
        
        ttk.Label(frame, text="DSP Tracker Launcher", font=("Arial", 12, "bold")).pack(pady=10)
        self.status_label = ttk.Label(frame, text="Checking for updates...", font=("Arial", 9))
        self.status_label.pack(pady=5)
        
        # Start update check in background
        threading.Thread(target=self.run_update_check, daemon=True).start()
        
        self.root.mainloop()

    def ensure_folders_exist(self):
        """Creates the data folders next to the .exe if they don't exist."""
        folders = ["Weekly_Scorecards", "Six_Week_Trailing"]
        for folder in folders:
            if not os.path.exists(folder):
                try:
                    os.makedirs(folder)
                    print(f"Created folder: {folder}")
                except Exception as e:
                    print(f"Could not create {folder}: {e}")

    def get_local_version(self):
        if not os.path.exists(SCRIPT_NAME):
            return "0.0.0"
        
        try:
            with open(SCRIPT_NAME, "r", encoding="utf-8") as f:
                content = f.read()
                # Regex to find VERSION = "x.x.x"
                match = re.search(r'VERSION\s*=\s*[\'"]([^\'"]+)[\'"]', content)
                if match:
                    return match.group(1)
        except Exception as e:
            print(f"Error reading local version: {e}")
        return "0.0.0"

    def get_remote_version(self, remote_content):
        try:
            match = re.search(r'VERSION\s*=\s*[\'"]([^\'"]+)[\'"]', remote_content)
            if match:
                return match.group(1)
        except Exception:
            pass
        return "0.0.0"

    def is_newer(self, remote_ver, local_ver):
        try:
            r_parts = [int(x) for x in remote_ver.split('.')]
            l_parts = [int(x) for x in local_ver.split('.')]
            return r_parts > l_parts
        except:
            return False

    def run_update_check(self):
        try:
            # 1. Check Local existence
            if not os.path.exists(SCRIPT_NAME):
                self.update_status("Downloading application...")
                self.do_update()
                self.launch_app()
                return

            # 2. Check Remote
            local_ver = self.get_local_version()
            self.update_status(f"Ver {local_ver}. Checking remote...")
            
            response = requests.get(GITHUB_RAW_URL, timeout=5)
            
            if response.status_code == 200:
                remote_content = response.text
                remote_ver = self.get_remote_version(remote_content)
                
                if self.is_newer(remote_ver, local_ver):
                    self.update_status(f"Updating to {remote_ver}...")
                    with open(SCRIPT_NAME, "w", encoding="utf-8") as f:
                        f.write(remote_content)
                    self.update_status("Update Complete.")
                else:
                    self.update_status("Up to date.")
            else:
                self.update_status("Offline mode.")

        except Exception as e:
            # CHANGE THIS BLOCK TEMPORARILY TO SEE THE ERROR
            print(f"Update error: {e}")
            self.update_status(f"Error: {str(e)[:40]}...") # Show first 40 chars of error in UI
            # messagebox.showerror("Update Failed", f"Detailed error:\n{e}") # Uncomment this for a popup
            
            # Wait a moment so you can read the error before launching local
            self.root.after(3000, self.launch_app)
            return 
        
        # 3. Launch
        self.root.after(1000, self.launch_app)

    def update_status(self, text):
        self.status_label.config(text=text)

    def launch_app(self):
        self.root.destroy()
        
        # Ensure the script exists
        if not os.path.exists(SCRIPT_NAME):
            messagebox.showerror("Error", f"Could not find {SCRIPT_NAME}")
            return

        try:
            # Dynamically import and run the external script
            spec = importlib.util.spec_from_file_location("scorecard_tracker", SCRIPT_NAME)
            module = importlib.util.module_from_spec(spec)
            sys.modules["scorecard_tracker"] = module
            spec.loader.exec_module(module)
            
            # Run the main function
            if hasattr(module, "main"):
                module.main()
            else:
                messagebox.showerror("Error", "The updated script does not have a 'main()' function.")
                
        except Exception as e:
            messagebox.showerror("Crash", f"Application crashed:\n{e}")

if __name__ == "__main__":
    # Ensure CWD is the executable folder (important for finding local files)
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
        os.chdir(application_path)
    
    Launcher()

# pyinstaller --noconsole --onefile --name="Scorecard Tracker App" --collect-all tksheet --collect-all tkcalendar --collect-all babel --collect-all pandas --collect-all numpy --collect-all openpyxl --collect-all tkinter launcher.py