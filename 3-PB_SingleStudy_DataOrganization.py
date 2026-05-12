import pandas as pd
import os
from pathlib import Path

# --- CONFIGURATION ---
# Path to the folder containing your Fz_Displacement_Analysis_...xlsx files
raw_data_root = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Study_Files"
# Path to your Master Excel where bone measurements (length/diameter) are stored
master_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Master_Measurements.xlsx"
output_file = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS_Study_Compiled.csv"

# Study Constants
STUDY_AGE = "16"
STUDY_SEX = "Female"

# Mapping the prefixes to your actual group names
GROUP_MAP = {
    "CV": "Control_Medigel",
    "PV": "IFS_Medigel",
    "PS": "IFS_SHP_Medigel"
}


def compile_ifs_study():
    # 1. Load Master Excel (assumes Mouse IDs are in the first column)
    # We load it once at the start to make lookups fast
    df_master = pd.read_excel(master_file)
    # Clean the IDs in the master sheet (remove spaces/ensure string)
    df_master.iloc[:, 0] = df_master.iloc[:, 0].astype(str).str.strip()

    compiled_data = []

    # 2. Iterate through analysis files
    root_path = Path(raw_data_root)
    # Find all analysis files in the folder and subfolders
    analysis_files = list(root_path.rglob("Fz_Displacement_Analysis_*.xlsx"))

    if not analysis_files:
        print("No analysis files found! Check your raw_data_root path.")
        return

    for file_path in analysis_files:
        # Extract ID from filename (e.g., "Fz_Displacement_Analysis_CV1_Femur.xlsx" -> "CV1")
        filename_parts = file_path.stem.replace(
            "Fz_Displacement_Analysis_", "").split('_')
        mouse_id = filename_parts[0].strip()

        # Determine Group from the prefix (CV, PV, PS)
        group_prefix = "".join([char for char in mouse_id if char.isalpha()])
        group_name = GROUP_MAP.get(group_prefix, "Unknown")

        # 3. Get Machine Data from the current file
        df_mach = pd.read_excel(file_path)

        # 4. Find matching row in Master Excel
        # Looks for the row where Column 0 matches the mouse_id (e.g., "CV1")
        master_row = df_master[df_master.iloc[:, 0] == mouse_id]

        if not master_row.empty:
            # Create a combined record
            record = {
                "Mouse_ID": mouse_id,
                "Group": group_name,
                "Age": STUDY_AGE,
                "Sex": STUDY_SEX,
                # Machine Metrics (from the .xlsx analysis file)
                "Max_Load_N": df_mach['Max_Load_N'].iloc[0],
                "Stiffness_N_mm": df_mach['Stiffness_N_per_mm'].iloc[0],
                "Work_to_Failure_Nmm": df_mach['Energy_to_Failure_Nmm'].iloc[0],
                "Disp_at_Failure_mm": df_mach['Displacement_at_Failure_mm'].iloc[0],
                # Measurement Metrics (from your Master Excel)
                # Adjust index [0, 1], [0, 2] etc. based on which columns hold your measurements
                "Length_mm": master_row.iloc[0, 1],
                "Diameter_mm": master_row.iloc[0, 2],
                "Thickness_mm": master_row.iloc[0, 3]
            }
            compiled_data.append(record)
            print(f"Successfully matched: {mouse_id}")
        else:
            print(f"Warning: Mouse ID '{mouse_id}' not found in Master Excel.")

    # 5. Save everything
    if compiled_data:
        final_df = pd.DataFrame(compiled_data)
        final_df.to_csv(output_file, index=False)
        print(f"\nSuccess! Compiled data saved to: {output_file}")
    else:
        print("\nNo data was compiled. Check your ID matching.")


if __name__ == "__main__":
    compile_ifs_study()
