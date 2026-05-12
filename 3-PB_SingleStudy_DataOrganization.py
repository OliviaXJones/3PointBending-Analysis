import os
import pandas as pd
import openpyxl
from pathlib import Path

# --- CONFIGURATION ---
# Root folder for machine analysis files
raw_data_root = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Study_Files"
# Your main Master Excel that you want to fill with data
master_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Master_Measurements.xlsx"
# The specific file where length/diameter/thickness are kept
measurement_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Measurement Files\IFS_Measurement_Data.xlsx"
# Final output for Prism/Analysis
output_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Study_Compiled.csv"

# Mapping prefixes to Group Names
GROUP_MAP = {
    "CV": "Control + Medigel",
    "PV": "IFS + Medigel",
    "PS": "IFS + SHP Medigel"
}


def sync_data_to_master():
    # 1. Load the Physical Measurements (Excel)
    print("Loading measurement files...")
    df_meas = pd.read_excel(measurement_file)
    # Clean ID column (assumed column 0)
    df_meas.iloc[:, 0] = df_meas.iloc[:, 0].astype(str).str.strip()

    # 2. Map the Machine Analysis Files
    print("Scanning for force-displacement analysis files...")
    root_path = Path(raw_data_root)
    analysis_files = list(root_path.rglob("Fz_Displacement_Analysis_*.xlsx"))

    # Map Mouse ID -> File Path for quick lookup
    file_map = {}
    for fp in analysis_files:
        # Extract ID (e.g., CV1) from "Fz_Displacement_Analysis_CV1_Femur.xlsx"
        mid = fp.stem.replace("Fz_Displacement_Analysis_",
                              "").split('_')[0].strip()
        file_map[mid] = fp

    # 3. Open Master Workbook for Writing
    wb = openpyxl.load_workbook(master_file)
    ws = wb.active  # Or wb["SheetName"] if you have a specific sheet

    print(f"Updating Master Sheet: {master_file}")

    # Iterate through rows in Master (assuming headers are in Row 1)
    for row in range(2, ws.max_row + 1):
        mouse_id = str(ws.cell(row=row, column=1).value).strip()
        if not mouse_id or mouse_id == "None":
            continue

        # --- Part A: Pull from Measurement File ---
        m_row = df_meas[df_meas.iloc[:, 0] == mouse_id]
        if not m_row.empty:
            # Adjust column numbers (2, 3, 4) to match your Master's layout
            ws.cell(row=row, column=1).value = m_row.iloc[0, 1]  # Length
            ws.cell(row=row, column=8).value = m_row.iloc[0, 2]  # Diameter
            ws.cell(row=row, column=12).value = m_row.iloc[0, 3]  # Thickness

        # --- Part B: Pull from Force-Displacement Analysis ---
        if mouse_id in file_map:
            df_mach = pd.read_excel(file_map[mouse_id])
            # Adjust column numbers to where you want these in your Master Excel
            ws.cell(row=row, column=4).value = df_mach['Max_Load_N'].iloc[0]
            ws.cell(
                row=row, column=5).value = df_mach['Stiffness_N_per_mm'].iloc[0]
            ws.cell(
                row=row, column=6).value = df_mach['Energy_to_Failure_Nmm'].iloc[0]
            ws.cell(
                row=row, column=7).value = df_mach['Displacement_at_Failure_mm'].iloc[0]
            print(f"Synced: {mouse_id}")
        else:
            print(f"Warning: No analysis file found for {mouse_id}")

    # 4. Save the Updated Master Excel
    wb.save(master_file)
    print("\nMaster Excel updated successfully.")

    # 5. Create the Compiled CSV for Analysis
    # Reloading the newly saved master to export a clean CSV
    final_df = pd.read_excel(master_file)
    final_df.to_csv(output_file, index=False)
    print(f"CSV Compiled: {output_file}")


if __name__ == "__main__":
    try:
        sync_data_to_master()
    except Exception as e:
        print(f"An error occurred: {e}")
