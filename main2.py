import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import requests
import json
import copy
import threading
import urllib3
import os
from datetime import datetime
import pyodbc
import warnings
import sys

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

warnings.filterwarnings('ignore') #warnings.filterwarnings('default')
stage_prod = 'phont57801ussql,14481' #Adjust the 57802(Stage) to 57801(Prod)


#BASE_URL = "https://ckb-floorplan-suite.fp-api.stage.space.walmartlabs.com/api/" #Stage
BASE_URL = "https://ckb-floorplan-suite.fp-api.space.walmartlabs.com/api/"
HEADERS = {
    'accept': 'application/json',
    'Content-Type': 'application/json',
    'WM_CONSUMER.ID': '412d9e23-a7bb-4ca1-8d68-f7174048c9b0', #03a3e299-129d-4b46-b7b2-1713c7808fcf
    'WM_SVC.NAME': 'SAD-FP-API',
    'WM_SVC.ENV': 'prod', #stg
}

default_payload = {
    "Id": 1, "FloorplanId": 9999, "PlanogramId": 9999, "DateModified": "2025-07-15T20:10:26.198Z", "WeekRange": None,
    "ExcludedWeekNumbers": None, "Changed": True, "Square": 0, "Cubic": 0, "LinearPercentage": 0, "SquarePercentage": 0,
    "LinearPercentageUsed": 0, "SquarePercentageUsed": 0, "Desc2": None, "Desc3": None, "TopStoreMovement": None,
    "BottomStoreMovement": None, "Desc6": None, "Desc7": None, "Desc8": None, "Desc9": None, "Desc10": None, "Desc11": None,
    "Desc12": None, "Desc13": None, "Desc14": None, "Desc15": None, "Desc16": None, "Desc17": None, "Desc18": None, "Desc19": None,
    "Desc20": None, "Desc21": None, "Desc22": None, "Desc23": None, "Desc24": None, "Desc25": None, "Desc26": None, "Desc27": None,
    "Desc28": None, "Desc29": None, "Desc30": None, "Desc31": None, "Desc32": None, "Desc33": None, "Desc34": None, "Desc35": None,
    "Desc36": None, "Desc37": None, "Desc38": None, "Desc39": None, "Desc40": None, "Desc41": None, "Desc42": None, "Desc43": None,
    "Desc44": None, "Desc45": None, "Desc47": None, "Desc48": None, "Desc49": None, "Desc50": None, "UnitMovement001Wk": 0,
    "UnitMovement004Wk": 0, "UnitMovement013Wk": 0, "UnitMovement026Wk": 0, "UnitMovement052Wk": 0, "UnitMovementFYTD": 0,
    "UnitMovementPFY": 0, "Price001Wk": 0, "Price004Wk": 0, "Price013Wk": 0, "Price026Wk": 0, "Price052Wk": 0, "PriceFYTD": 0,
    "PricePFY": 0, "UnitCost001Wk": 0, "UnitCost004Wk": 0, "UnitCost013Wk": 0, "UnitCost026Wk": 0, "UnitCost052Wk": 0,
    "UnitCostFYTD": 0, "UnitCostPFY": 0, "OnHand001Wk": 0, "OnHand004Wk": 0, "OnHand013Wk": 0, "OnHand026Wk": 0, "OnHand052Wk": 0,
    "OnHandFYTD": 0, "OnHandPFY": 0, "InstockPct001Wk": 0, "InstockPct004Wk": 0, "InstockPct013Wk": 0, "InstockPct026Wk": 0,
    "InstockPct052Wk": 0, "InstockPctFYTD": 0, "InstockPctPFY": 0, "UnitMovement": 0, "UnitCost": 0, "Price": 0, "OnHand": 0,
    "InstockPct": 0, "TotalSales": 0, "TotalMovement": 0, "Linear": 0, "Value43": None, "Value44": None, "Value45": None,
    "Value46": None, "Value47": None, "Value48": None, "Value49": None, "Value50": None, "Flag1": None, "Flag2": None,
    "Flag3": None, "Flag4": None, "Flag5": None, "Flag6": None, "Flag7": None, "Flag8": None, "Flag9": None, "Flag10": None,
    "PartId": None, "GLN": None, "Warning": None, "WarningNumber": 0, "TrafficFlow": None, "NumberOfFixtures": None, "NumberOfSections": 0,
    "CostAllocated": None, "NumberOfProductsAllocated": None, "ProfitAllocated": None, "RoiiCostAllocated": None, "RoiiRetailAllocated": None,
    "SalesAllocated": None, "AnnualProfitAllocated": None, "CombinedPerfIndexAllocated": None, "MarginAllocated": None,
    "MovementAllocated": None, "Capacity": None, "CapacityCost": None, "CapacityRetail": None, "CapacityUnrestricted": None,
    "AllocationTargetSpace": None, "MovementPeriodUsed": None, "DBDateEffectiveFrom": None, "DBDateEffectiveTo": None,
    "DBStatus": 0, "Active": None
}

# Trimmed default_payload for brevity, add other fields if needed.

class UploadApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Add Planogram to Floorplans")
        self.center_window()

        self.file_path = ""
        self.df = None

        # UI Components
        self.show_errors_only = tk.BooleanVar(value=False)
        self.create_widgets()
        self.center_window()
        self.valid_floorplans = self.query_floorplans()
        self.valid_planograms = self.query_planograms()


    def center_window(self, width=1000, height=700):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        # Create a horizontal frame to hold all buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        # ⬅️ Add Download Template button first (leftmost)
        tk.Button(button_frame, text="Download Template", command=self.download_template).pack(side=tk.LEFT, padx=5)

        tk.Button(button_frame, text="Select CSV", command=self.load_csv).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Start Upload", command=self.start_upload).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Exit", command=self.root.destroy).pack(side=tk.LEFT, padx=5)

        # Treeview (table)
        self.tree = ttk.Treeview(self.root)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree.tag_configure("current_upload", background="#c6f5c6")

        # Progress bar
        self.progress = ttk.Progressbar(self.root, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, padx=10, pady=5)

        tk.Checkbutton(
            button_frame,
            text="Show Errors Only",
            variable=self.show_errors_only,
            command=self.toggle_error_filter
        ).pack(side=tk.LEFT, padx=5)


    def download_template(self):
        from tkinter import filedialog
        import pandas as pd

        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Template As"
        )
        if save_path:
            template_df = pd.DataFrame(columns=["floorplan_dbkey", "planogram_dbkey"])
            template_df.to_csv(save_path, index=False)
            messagebox.showinfo("Template Downloaded", f"Template saved to:\n{save_path}")

    def load_csv(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if self.file_path:
            self.df = pd.read_csv(self.file_path)
            self.df["status"] = ""  # To show ✔️ or ❌
            self.df["message"] = ""  # To show response
            self.display_dataframe()

    def toggle_error_filter(self):
        self.display_dataframe()

    def display_dataframe(self, highlight_rows=None):
        highlight_rows = highlight_rows or []

        for widget in self.tree.get_children():
            self.tree.delete(widget)

        display_df = self.df.copy()

        # Filter if checkbox is checked
        if self.show_errors_only.get():
            display_df = display_df[display_df['status'] == '❌']

        columns = list(display_df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", stretch=True, width=150)

        for idx, row in display_df.iterrows():
            tags = ("current_upload",) if idx in highlight_rows else ()
            self.tree.insert("", "end", values=list(row), tags=tags)

    def start_upload(self):
        if self.df is not None:
            thread = threading.Thread(target=self.upload_data)
            thread.start()
        else:
            self.output.insert(tk.END, "❗ Please load a CSV file first.\n")

    def upload_data(self):
        grouped = self.df.groupby('floorplan_dbkey')
        total = len(grouped)
        self.progress["maximum"] = total

        valid_floorplan_set = set(self.valid_floorplans['DBKEY'].astype(int))
        valid_planogram_set = set(self.valid_planograms['DBKEY'].astype(int))

        for i, (floorplan_id, group_df) in enumerate(grouped, 1):
            index_list = group_df.index.tolist()

            # Floorplan check
            if int(floorplan_id) not in valid_floorplan_set:
                self.df.loc[index_list, 'status'] = '❌'
                self.df.loc[index_list, 'message'] = "No WIP floorplan with this DBKEY found."
                self.progress["value"] = i
                self.display_dataframe(highlight_rows=index_list)
                self.tree.see(self.tree.get_children()[index_list[0]])
                continue

            payload_list = []
            bad_indexes = []

            for idx, row in group_df.iterrows():
                planogram_id = int(row['planogram_dbkey'])

                if planogram_id not in valid_planogram_set:
                    self.df.at[idx, 'status'] = '❌'
                    self.df.at[idx, 'message'] = "No LIVE Planogram found for this DBKEY."
                    bad_indexes.append(idx)
                    continue

                payload = copy.deepcopy(default_payload)
                payload["FloorplanId"] = int(row['floorplan_dbkey'])
                payload["PlanogramId"] = planogram_id
                payload_list.append(payload)

            # If all records were invalid, skip the POST
            if not payload_list:
                self.progress["value"] = i
                self.display_dataframe(highlight_rows=index_list)
                self.tree.see(self.tree.get_children()[index_list[0]])
                continue

            url = f"{BASE_URL}floorplans/{floorplan_id}/performance/batch"
            response = requests.post(
                url,
                headers=HEADERS,
                data=json.dumps(payload_list),
                verify=False
            )

            status = "✔️" if response.status_code in [200, 201] else "❌"
            message = "Success" if status == "✔️" else f"{response.status_code}: {response.text.strip()}"

            # Mark only valid records as successful or failed from POST
            for idx in group_df.index:
                if idx not in bad_indexes:
                    self.df.at[idx, 'status'] = status
                    self.df.at[idx, 'message'] = message

            self.progress["value"] = i
            self.display_dataframe(highlight_rows=index_list)
            self.tree.see(self.tree.get_children()[index_list[0]])

        self.display_dataframe()  # Final refresh to clear highlights
        self.save_results_to_csv(self.df)

    def save_results_to_csv(self, final_df):
        base_path = os.path.dirname(self.file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result_file = os.path.join(base_path, f"upload_results_{timestamp}.csv")

        # 🔁 Convert display symbols to plain text for export
        export_df = final_df.copy()
        export_df["status"] = export_df["status"].map({
            "✔️": "Success",
            "❌": "Unsuccessful"
        }).fillna(export_df["status"])  # In case any other values exist

        export_df.to_csv(result_file, index=False)

        messagebox.showinfo("Export Complete", f"✅ Upload results saved to:\n\n{result_file}")

    def query_floorplans(self):
        try:
            conn_ckb = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                f'SERVER={stage_prod};'
                'DATABASE=us_ckb;'
                'Trusted_Connection=yes;'
            )
            query = "SELECT DBKEY FROM IX_FLR_FLOORPLAN WHERE DBSTATUS = 3"
            df = pd.read_sql(query, conn_ckb)
            conn_ckb.close()
            return df
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to query floorplans:\n{e}")
            return pd.DataFrame()

    def query_planograms(self):
        try:
            conn_ckb = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                f'SERVER={stage_prod};'
                'DATABASE=us_ckb;'
                'Trusted_Connection=yes;'
            )
            query = """
                SELECT DBKEY FROM IX_SPC_PLANOGRAM
                WHERE DBSTATUS = 1
                AND DESC39 IN ('SPACEPPOG','NMSPACE')
            """
            df = pd.read_sql(query, conn_ckb)
            conn_ckb.close()
            return df
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to query planograms:\n{e}")
            return pd.DataFrame()

if __name__ == "__main__":
    root = tk.Tk()
    app = UploadApp(root)
    root.mainloop()
