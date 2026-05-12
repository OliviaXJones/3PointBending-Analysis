import os
import pandas as pd
import openpyxl
from pathlib import Path

# --- CONFIGURATION ---
# Root folder for machine analysis files
raw_data_root = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Force-Displacement Raw Files\IFS+SHP099+Medigel_LFemur_051226"
# Your main Master Excel that you want to fill with data
master_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS+SHP099+Medigel_LFemurMaster.xlsx"
# The specific file where length/diameter/thickness are kept
measurement_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Measurement Files\IFS+SHP099+Medigel_LFemur_051226.xlsx"
# Final output for Prism/Analysis
output_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Study_Compiled.csv"

# Mapping prefixes to Group Names
GROUP_MAP = {
    "CV": "Control + Medigel",
    "PV": "IFS + Medigel",
    "PS": "IFS + SHP Medigel"
}


def get_analysis_file(folder_path):
    root = Path(folder_path)
    # Search for any .xlsx file starting with the analysis prefix
    files = list(root.glob("Fz_Displacement_Analysis_*.xlsx"))

    if not files:
        raise FileNotFoundError(
            f"Could not find an analysis file in {folder_path}")

    # If there are multiple (e.g., from different dates), pick the most recent one
    return max(files, key=os.path.getmtime)


def sync_data_to_master():
    # 1. Load the Physical Measurements
    print("Loading measurement files...")
    df_meas = pd.read_excel(measurement_file)
    df_meas.iloc[:, 0] = df_meas.iloc[:, 0].astype(str).str.strip()

 # 2. Dynamically find and load the single Analysis File
    analysis_file_path = get_analysis_file(raw_data_root)
    df_all_mach = pd.read_excel(analysis_file_path)

    # Instead of using the name 'Mouse Code', we use .iloc[:, 0]
    # This means: "Take the very first column, no matter what it is named"
    df_all_mach['Mouse Code'] = df_all_mach.iloc[:, 0].astype(
        str).str.replace('.txt', '', regex=False).str.strip()

    # Clean the ID column in the analysis file (adjust name if it's not 'Mouse Code')
    id_col_name = 'Filename'
    df_all_mach[id_col_name] = df_all_mach[id_col_name].astype(str).str.strip()

    # 3. Open Master Workbook
    wb = openpyxl.load_workbook(master_file)
    ws = wb.active

    for row in range(2, ws.max_row + 1):
        # 1. Grab the ID from Column A (1) BEFORE doing anything else
        mouse_id_cell = ws.cell(row=row, column=1).value
        if mouse_id_cell is None:
            continue

        mouse_id = str(mouse_id_cell).strip()

        # --- Part A: Pull from Measurement File ---
        m_row = df_meas[df_meas.iloc[:, 0] == mouse_id]
        if not m_row.empty:
            # Source: Index 1 (Length) -> Master: Column 2 (AvgLength)
            ws.cell(row=row, column=2).value = m_row.iloc[0, 1]

            # Source: Index 8 (D_avg) -> Master: Column 3 (AvgDiam)
            ws.cell(row=row, column=3).value = m_row.iloc[0, 8]

            # Source: Index 12 (h_avg) -> Master: Column 4 (AvgThick)
            ws.cell(row=row, column=4).value = m_row.iloc[0, 12]

        # --- Part B: Pull from Force-Displacement Analysis ---
        mach_row = df_all_mach[df_all_mach['Mouse Code'] == mouse_id]

        if not mach_row.empty:
            # Master columns 5, 6, 7, 8 match your Max Load through Displacement headers
            ws.cell(row=row, column=5).value = mach_row.iloc[0]['Max_Load_N']
            ws.cell(
                row=row, column=6).value = mach_row.iloc[0]['Stiffness_N_per_mm']
            ws.cell(
                row=row, column=7).value = mach_row.iloc[0]['Energy_to_Failure_Nmm']
            ws.cell(
                row=row, column=8).value = mach_row.iloc[0]['Displacement_at_Failure_mm']
            print(f"Synced: {mouse_id}")
        else:
            print(f"Warning: {mouse_id} not found in analysis file.")

    wb.save(master_file)
    print("\nMaster Excel updated successfully.")


if __name__ == "__main__":
    try:
        sync_data_to_master()
    except Exception as e:
        print(f"An error occurred: {e}")
