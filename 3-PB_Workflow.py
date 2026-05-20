import glob
from datetime import datetime
from io import StringIO
import os
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd

# ===========================================================
# CONFIGURATION & CONSTANTS
# ===========================================================
BASE_DIR = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending"
RAW_DATA_ROOT = os.path.join(
    BASE_DIR, "Force-Displacement Raw Files", "FKBP5Null_Tibia_11226"
)
MASTER_FILE = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\FKBP5_3-PointBendingTibiaMaster.xlsx"
MEASUREMENT_FILE = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Measurement Files\FKBP5Null_Tibia+Femur_11226.xlsx"
OUTPUT_FILE = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\FKBP5_3-PointBendingTibiaCompiled.csv"

FILE_GLOB_PATTERN = "*.txt"
SAVE_PNG_DPI = 100
TOE_LOAD_FRACTION = 0.05
LINEAR_WINDOW_POINTS = 90
MIN_R2 = 0.995

# ===========================================================
# PART 1: BENDING DATA ANALYZER FUNCTIONS
# ===========================================================


def read_bending_txt(filepath):
    with open(filepath, "r", errors="ignore") as f:
        lines = f.readlines()

    data_start = None
    data_end = None
    for i, line in enumerate(lines):
        if "<DATA>" in line:
            data_start = i + 1
        elif "<END DATA>" in line and data_start:
            data_end = i
            break
    if data_start is None or data_end is None:
        raise ValueError(f"No valid <DATA> section in {filepath}")

    data_lines = lines[data_start:data_end]
    header = None
    for i, line in enumerate(data_lines):
        parts = line.strip().split("\t")
        try:
            float(parts[0])
        except ValueError:
            header = parts
            data_lines = data_lines[i + 1:]
            break

    df = pd.read_csv(
        StringIO("".join(data_lines)), sep="\t", names=header, engine="python"
    )
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["Fz, N", "Position (z), mm"], how="any")
    if "Fx" in df.columns:
        df["Fz, N"] = df["Fx"]
    if "Position (z)" in df.columns:
        df["Position (z), mm"] = df["Position (z)"]
    df["Fz, N"] = -df["Fz, N"]
    return df


def dominant_linear_region(x, y, window=30, min_r2=0.995):
    best_len = 0
    best = (0, 0, 0, 0)  # m, b, start, end
    n = len(x)

    if n < window:
        return 0, 0, 0, n

    for start in range(0, n - window, 2):
        end = start + window
        x_seg = x[start:end]
        y_seg = y[start:end]

        m, b = np.polyfit(x_seg, y_seg, 1)
        r = np.corrcoef(x_seg, y_seg)[0, 1]
        r2 = r**2

        if r2 >= min_r2:
            while end < n:
                next_end = min(end + 5, n)
                x_ext = x[start:next_end]
                y_ext = y[start:next_end]
                m_ext, b_ext = np.polyfit(x_ext, y_ext, 1)
                r_ext = np.corrcoef(x_ext, y_ext)[0, 1]
                if (r_ext**2) < min_r2:
                    break
                m, b, end = m_ext, b_ext, next_end

            if (end - start) > best_len:
                best_len = end - start
                best = (m, b, start, end)

    return best


def run_batch_bending_analysis(input_folder):
    """Analyzes raw force-displacement curves within a specific directory."""
    print(f"\n--- Starting Batch Bending Analysis in: {input_folder} ---")
    date_str = datetime.now().strftime("%m%d%y")

    plot_folder = os.path.join(input_folder, "Fz_Displacement_Analysis")
    os.makedirs(plot_folder, exist_ok=True)

    excel_filename = f"Fz_Displacement_Analysis_{date_str}.xlsx"

    txt_files = glob.glob(os.path.join(input_folder, FILE_GLOB_PATTERN))
    if not txt_files:
        print(f"No TXT files found in {input_folder}")
        return

    results = []

    for path in txt_files:
        try:
            df = read_bending_txt(path)
            displacement = df["Position (z), mm"].values
            load = df["Fz, N"].values

            load_smooth = np.convolve(load, np.ones(3) / 3, mode="same")

            DISPLACEMENT_LIMIT = 1.75
            valid_range_mask = displacement <= DISPLACEMENT_LIMIT

            load_for_peak = load_smooth[valid_range_mask]
            constrained_max = (
                np.max(load_for_peak)
                if len(load_for_peak) > 0
                else np.max(load_smooth)
            )

            candidates = np.where(
                (load_smooth >= 0.5 * constrained_max) & (valid_range_mask)
            )[0]
            max_idx = (
                candidates[np.argmax(load[candidates])]
                if len(candidates) > 0
                else np.argmax(load_smooth)
            )

            max_load = load[max_idx]
            disp_at_max = displacement[max_idx]

            post_max_load = load[max_idx:]
            NEAR_ZERO_THRESHOLD = 0.5
            drop_threshold_value = max_load * 0.80

            zero_drop_indices = np.where(post_max_load <= NEAR_ZERO_THRESHOLD)[
                0
            ]
            if len(zero_drop_indices) > 0:
                fail_idx = max_idx + zero_drop_indices[0]
            else:
                has_dropped_indices = np.where(
                    post_max_load < drop_threshold_value
                )[0]
                if len(has_dropped_indices) > 0:
                    search_start = has_dropped_indices[0]
                    search_load = post_max_load[search_start:]
                    diff_post = np.diff(
                        np.convolve(search_load, np.ones(3) / 3, mode="same")
                    )
                    local_min_indices = np.where(diff_post >= 0)[0]
                    fail_idx = (
                        max_idx
                        + search_start
                        + (
                            local_min_indices[0]
                            if len(local_min_indices) > 0
                            else len(search_load) - 1
                        )
                    )
                else:
                    fail_idx = len(load) - 1

            pre_max_disp = displacement[:max_idx]
            pre_max_load = load[:max_idx]

            toe_mask = pre_max_load >= (TOE_LOAD_FRACTION * max_load)
            toe_indices = np.where(toe_mask)[0]

            start_idx = toe_indices[0] if len(toe_indices) > 0 else 0
            adj_disp = displacement - displacement[start_idx]

            adj_max_disp = adj_disp[max_idx]
            adj_fail_disp = adj_disp[fail_idx]

            disp_at_failure = adj_fail_disp
            load_at_failure = load[fail_idx]

            disp_slope_candidates = (
                pre_max_disp[toe_mask] - displacement[start_idx]
            )
            load_slope_candidates = pre_max_load[toe_mask]

            if len(disp_slope_candidates) < LINEAR_WINDOW_POINTS:
                print(
                    f"Skipping stiffness fit (too few points) in {os.path.basename(path)}"
                )
                stiffness = np.nan
                intercept = np.nan
                idx0, idx1 = 0, 0
            else:
                stiffness, intercept, idx0, idx1 = dominant_linear_region(
                    disp_slope_candidates,
                    load_slope_candidates,
                    window=LINEAR_WINDOW_POINTS,
                )

            energy = np.trapezoid(
                load[start_idx: fail_idx + 1], adj_disp[start_idx: fail_idx + 1]
            )

            # Individual curve generation
            plt.figure(figsize=(7, 5))
            plt.plot(adj_disp, load, color="black", label="Data")
            if (idx1 - idx0) > 5 and np.isfinite(stiffness):
                x_lin = disp_slope_candidates[idx0:idx1]
                plt.plot(
                    x_lin,
                    stiffness * x_lin + intercept,
                    color="red",
                    label=f"Stiffness: {stiffness:.2f} N/mm",
                )

            plt.scatter(
                adj_max_disp, max_load, color="green", label="Max Load"
            )
            plt.scatter(
                adj_fail_disp,
                load_at_failure,
                color="purple",
                label="Failure",
            )
            plt.axvline(
                x=0, color="blue", linestyle="--", label="Toe End (Start)"
            )
            plt.fill_between(
                adj_disp[start_idx: fail_idx + 1],
                load[start_idx: fail_idx + 1],
                alpha=0.2,
                color="orange",
            )
            plt.title(os.path.basename(path))
            plt.legend()
            plt.savefig(
                os.path.join(
                    plot_folder, os.path.basename(path).replace(".txt", ".png")
                ),
                dpi=SAVE_PNG_DPI,
            )
            plt.close()

            results.append(
                {
                    "Filename": os.path.basename(path),
                    "Max_Load_N": round(max_load, 4),
                    "Stiffness_N_per_mm": round(stiffness, 4),
                    "Energy_to_Failure_Nmm": round(energy, 4),
                    "Displacement_at_Failure_mm": round(disp_at_failure, 4),
                }
            )
            print(f"Processed: {os.path.basename(path)}")

        except Exception as e:
            print(f"Error {os.path.basename(path)}: {e}")

    output_excel_path = os.path.join(input_folder, excel_filename)
    pd.DataFrame(results).to_excel(output_excel_path, index=False)
    print(f"Saved analysis summaries to: {output_excel_path}")


# ===========================================================
# PART 2: MASTER MERGING & EXPORT FUNCTIONS
# ===========================================================


def sync_data_to_master(all_analysis_files, master_file, measurement_file):
    df_meas_all = pd.read_excel(measurement_file, sheet_name=None)
    wb = openpyxl.load_workbook(master_file)
    folder_to_file_map = {}

    def get_date(path_obj):
        try:
            return datetime.strptime(path_obj.stem.split("_")[-1], "%m%d%y")
        except ValueError:
            return datetime.min

    for file_path in all_analysis_files:
        folder_name = file_path.parent.name
        if folder_name not in folder_to_file_map:
            folder_to_file_map[folder_name] = file_path
        else:
            if get_date(file_path) > get_date(folder_to_file_map[folder_name]):
                folder_to_file_map[folder_name] = file_path

    for sheet_name in wb.sheetnames:
        if sheet_name.lower() in ["summary", "notes", "calculations"]:
            continue

        if sheet_name not in folder_to_file_map:
            print(f"Skipping {sheet_name}: No matching folder found.")
            continue

        print(
            f"Syncing Sheet '{sheet_name}' using file: {folder_to_file_map[sheet_name].name}"
        )

        ws = wb[sheet_name]
        df_mach = pd.read_excel(folder_to_file_map[sheet_name])
        df_mach["Mouse Code"] = df_mach["Filename"].str.replace(
            ".txt", "", regex=False
        )
        df_meas = df_meas_all.get(sheet_name, pd.DataFrame())

        for row in range(2, ws.max_row + 1):
            for start_col in [1, 10]:  # Column A (Males) and J (Females)
                mouse_code = ws.cell(row=row, column=start_col).value
                if not mouse_code:
                    continue

                mouse_code = str(mouse_code).strip()

                if not df_meas.empty:
                    m_row = df_meas[
                        df_meas.iloc[:, 0].astype(
                            str).str.strip() == mouse_code
                    ]
                    if not m_row.empty:
                        ws.cell(row=row, column=start_col + 1).value = (
                            m_row.iloc[0, 1]
                        )
                        ws.cell(row=row, column=start_col + 2).value = (
                            m_row.iloc[0, 8]
                        )
                        ws.cell(row=row, column=start_col + 3).value = (
                            m_row.iloc[0, 12]
                        )

                mach_row = df_mach[
                    df_mach["Mouse Code"].astype(str).str.strip() == mouse_code
                ]
                if not mach_row.empty:
                    ws.cell(row=row, column=start_col + 4).value = (
                        mach_row.iloc[0]["Max_Load_N"]
                    )
                    ws.cell(row=row, column=start_col + 5).value = (
                        mach_row.iloc[0]["Stiffness_N_per_mm"]
                    )
                    ws.cell(row=row, column=start_col + 6).value = (
                        mach_row.iloc[0]["Energy_to_Failure_Nmm"]
                    )
                    ws.cell(row=row, column=start_col + 7).value = (
                        mach_row.iloc[0]["Displacement_at_Failure_mm"]
                    )

    wb.save(master_file)
    print("Master file update complete.")


def parse_mouse_code(code):
    try:
        parts = str(code).split(".")
        genotype = parts[0]
        age = parts[1]
        sex_id = parts[2]
        sex = "Male" if sex_id.startswith("M") else "Female"
        mouse_num = sex_id[1:]
        return genotype, age, sex, mouse_num
    except IndexError:
        return None, None, None, None


def generate_grouped_tables(df):
    metrics = [
        "Avg. Tibia Length",
        "Avg. Tibia Diameter",
        "Avg. Tibia Thickness",
        "Maximum Load",
        "Stiffness",
        "Energy to Failure",
        "Displacement at Failure",
    ]
    unique_ages = df["Age_Extracted"].unique()
    sort_priority = ["Wildtype", "Mutant", "Heterozygous"]

    folder_genotype = os.path.join(BASE_DIR, "Tibia_Analysis_By_Genotype")
    folder_lineage = os.path.join(BASE_DIR, "Tibia_Analysis_By_Lineage")
    os.makedirs(folder_genotype, exist_ok=True)
    os.makedirs(folder_lineage, exist_ok=True)

    for metric in metrics:
        for age in unique_ages:
            for sex in ["Male", "Female"]:
                base_subset = df[
                    (df["Sex_Extracted"] == sex)
                    & (df["Age_Extracted"] == age)
                ].copy()
                if base_subset.empty:
                    continue

                clean_metric = metric.replace("_", " ").replace(".", "")

                # Set 1: Analysis by Genotype
                gen_subset = base_subset[
                    base_subset["Progeny_Group"].isin(sort_priority)
                ].copy()
                if not gen_subset.empty:
                    table_genotype = gen_subset.pivot_table(
                        index="ID_Num", columns="Progeny_Group", values=metric
                    )
                    table_genotype = table_genotype.reindex(
                        columns=sort_priority
                    )
                    name_genotype = f"{sex} {age}Wks {clean_metric}.csv"
                    table_genotype.to_csv(
                        os.path.join(folder_genotype, name_genotype)
                    )

                # Set 2: Analysis by Lineage
                table_lineage = base_subset.pivot_table(
                    index="ID_Num", columns="Progeny_Group", values=metric
                )

                def lineage_sort(col_name):
                    for i, gen in enumerate(sort_priority):
                        if col_name.startswith(gen):
                            return (i, col_name)
                    return (99, col_name)

                sorted_lineage_cols = sorted(
                    table_lineage.columns, key=lineage_sort
                )
                table_lineage = table_lineage[sorted_lineage_cols]
                name_lineage = f"{sex} {age}Wks {clean_metric}.csv"
                table_lineage.to_csv(os.path.join(
                    folder_lineage, name_lineage))


def process_all_sheets(raw_df, structure_type="Single Table"):
    """
    Processes the master sheet based on the user-selected data structure.
    structure_type can be "Split Genders (Male/Female)" or "Single Table"
    """
    # Clean up empty rows
    raw_df = raw_df.dropna(how='all')

    if structure_type == "Split Genders (Male/Female)":
        # Assumes Column I (index 8) is a blank spacer
        males = raw_df.iloc[:, 0:8].copy()
        females = raw_df.iloc[:, 9:17].copy()

        # Set proper headers for both
        males.columns = males.iloc[0]
        males = males[1:].dropna(how='all')
        males['Gender'] = 'Male'

        females.columns = females.iloc[0]
        females = females[1:].dropna(how='all')
        females['Gender'] = 'Female'

        # Combine them vertically into a single clean dataset
        combined_df = pd.concat([males, females], ignore_index=True)
        return combined_df

    else:
        # "Single Table" mode: Use the columns exactly as they are
        raw_df.columns = raw_df.iloc[0]
        combined_df = raw_df[1:].dropna(how='all')

        # If there isn't a Gender column, add a default one so downstream code doesn't break
        if 'Gender' not in combined_df.columns:
            combined_df['Gender'] = 'Combined/All'

        return combined_df


# ===========================================================
# PIPELINE EXECUTION
# ===========================================================
if __name__ == "__main__":
    try:
        # Step 1: Automatic batch processing of subdirectories
        print("Starting comprehensive pipeline...")
        root_path = Path(RAW_DATA_ROOT)

        # Find all distinct sub-folders that contain raw .txt data files
        subfolders_with_data = set()
        for txt_file in root_path.rglob(FILE_GLOB_PATTERN):
            subfolders_with_data.add(txt_file.parent)

        if not subfolders_with_data:
            print(
                f"No raw data .txt files found anywhere under paths matching: {RAW_DATA_ROOT}"
            )
            # Use manual fall-back window if automated lookup yielded nothing
            root = Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askdirectory(
                title="Select Raw Force-Displacement Folder"
            )
            root.destroy()
            if selected:
                subfolders_with_data.add(Path(selected))

        # Run the primary raw curve analysis for every folder discovered
        for folder in subfolders_with_data:
            run_batch_bending_analysis(str(folder))

        # Step 2: Dynamically aggregate the newly created excel summaries
        all_analysis_files = list(
            root_path.rglob("Fz_Displacement_Analysis_*.xlsx")
        )

        # Step 3: Run Master Sheets Merging, Synchronization, and Cross-Filtering
        sync_data_to_master(all_analysis_files, MASTER_FILE, MEASUREMENT_FILE)
        process_all_sheets(MASTER_FILE)

        print("\nAll pipeline tasks completed successfully!")

    except Exception as e:
        print(f"\nAn error halted execution of the pipeline: {e}")
