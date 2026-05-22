import os
import pandas as pd
import openpyxl
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, simpledialog

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


def get_save_location():
    """
    Opens windows to pick the parent folder AND type the new folder name.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes('-topmost', True)  # Bring to front

    # 1. Pop up for the Directory
    parent_dir = filedialog.askdirectory(
        title="Select where to save the results")

    if not parent_dir:
        return None

    # 2. Pop up for the Folder Name
    # This creates a small text box window so you don't have to use the terminal
    new_folder_name = simpledialog.askstring("Folder Name",
                                             "What should we name the new results folder?",
                                             initialvalue="LFemur_Results")

    if not new_folder_name:
        new_folder_name = "Grouped_Analysis_Results"

    full_path = os.path.join(parent_dir, new_folder_name)
    return full_path


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


def parse_ifs_code(code):
    """
    Parses IDs like 'CV1', 'PV10', 'PS5' into Group and Number.
    """
    try:
        # Extract letters (Group) and numbers (ID)
        import re
        match = re.match(r"([A-Z]+)([0-9]+)", str(code).strip())
        if match:
            prefix = match.group(1)
            num = match.group(2)
            group_name = GROUP_MAP.get(prefix, "Unknown")
            return group_name, prefix, num
    except:
        pass
    return "Unknown", "Unknown", "Unknown"


def export_final_tables(master_path):
    """
    Reads the updated Master Excel and exports Pivot Tables to a user-defined folder.
    """
    # 1. Ask user where to save and what to name it
    output_folder = get_save_location()
    if not output_folder:
        output_folder = os.path.join(os.path.dirname(
            master_path), "Grouped_Analysis_Tables")

    print(f"Generating grouped analysis tables in: {output_folder}")
    df = pd.read_excel(master_path)

    # Parse IDs
    df[['Group', 'Prefix', 'ID_Num']] = df.iloc[:, 0].apply(
        lambda x: pd.Series(parse_ifs_code(x)))

    metrics = {
        'AvgLength (mm)': 'Avg_Length',
        'AvgDiam (mm)': 'Avg_Diameter',
        'AvgThick (mm)': 'Avg_Thickness',
        'Max Load (N)': 'Max_Load',
        'Stiffness (N/mm)': 'Stiffness',
        'Energy to Failure (N*mm)': 'Energy_to_Failure',
        'Displacement at Failure (mm)': 'Displacement_at_Failure'
    }

    os.makedirs(output_folder, exist_ok=True)

    for header, clean_name in metrics.items():
        if header in df.columns:
            table = df.pivot_table(
                index='ID_Num', columns='Group', values=header)
            order = [GROUP_MAP[k]
                     for k in ["CV", "PV", "PS"] if GROUP_MAP[k] in table.columns]
            table = table.reindex(columns=order)

            file_name = f"{clean_name}.csv"
            table.to_csv(os.path.join(output_folder, file_name))
            print(f"  - Exported: {file_name}")


if __name__ == "__main__":
    try:
        # 1. Sync data from measurements and machine files to Master
        sync_data_to_master()

        # 2. Export the grouped tables with User Input for location/name
        export_final_tables(master_file)

        # 3. Final Master CSV
        final_df = pd.read_excel(master_file)
        final_df.to_csv(output_file, index=False)

        print("\nWorkflow Complete!")
    except Exception as e:
        print(f"\nAn error occurred: {e}")
