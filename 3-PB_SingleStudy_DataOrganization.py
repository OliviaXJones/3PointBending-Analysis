import os
import re
import glob
import json
from io import StringIO
from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd
import openpyxl
import matplotlib
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

matplotlib.use("Agg")

# ===========================================================
# STUDY CONFIGURATION & AUTOMATIC STATE MANAGER
# ===========================================================
CONFIG_JSON_FILENAME = "studies_config.json"

DEFAULT_FALLBACK_CONFIG = {
    "raw_data_root": r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Force-Displacement Raw Files\IFS+SHP099+Medigel_LFemur_051226",
    "master_file": r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS+SHP099+Medigel_LFemurMaster.xlsx",
    "measurement_file": r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Measurement Files\IFS+SHP099+Medigel_LFemur_051226.xlsx",
    "output_folder": r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\IFS+SHP99+Medigel 2026",
    "group_map": {
        "CV": "Control + Medigel",
        "PV": "IFS + Medigel",
        "PS": "IFS + SHP Medigel"
    }
}


def load_study_config():
    """Reads paths dynamically from your local config JSON if it exists."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_path = os.path.join(script_dir, CONFIG_JSON_FILENAME)
    if os.path.exists(json_path):
        try:
            with open(json_path, "r") as f:
                data = json.load(f)
                if "IFS+SHP099+Medigel 2026" in data:
                    return data["IFS+SHP099+Medigel 2026"]
                return data
        except Exception as e:
            print(f"Using default presets. State JSON notice: {e}")
    return DEFAULT_FALLBACK_CONFIG


def save_study_config(raw_dir, master_path, meas_path, out_dir):
    """Saves current directory selections so they are remembered next launch."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_path = os.path.join(script_dir, CONFIG_JSON_FILENAME)
    try:
        payload = {
            "IFS+SHP099+Medigel 2026": {
                "raw_data_root": raw_dir,
                "master_file": master_path,
                "measurement_file": meas_path,
                "output_folder": out_dir,
                "group_map": STUDY_CONFIG["group_map"]
            }
        }
        with open(json_path, "w") as f:
            json.dump(payload, f, indent=4)
    except Exception as e:
        print(f"Could not automatically update state file: {e}")


# Initial load of configuration parameters
STUDY_CONFIG = load_study_config()

FILE_GLOB_PATTERN = "*.txt"
SAVE_PNG_DPI = 100
TOE_LOAD_FRACTION = 0.05
LINEAR_WINDOW_POINTS = 90
MIN_R2 = 0.995


# ===========================================================
# ID PARSING UTIL (ROBUST GROUP + DIGIT EXTRACTOR)
# ===========================================================

def parse_id_and_group(text_string, group_map):
    """
    Looks for group prefixes followed by digits anywhere inside a piece of text.
    Returns: (clean_prefix_string, numerical_id_string, group_name) or (None, None, None)
    Example: 'CV12_LeftFemur.txt' -> ('CV', '12', 'Control + Medigel')
    """
    if not text_string or pd.isna(text_string):
        return None, None, None

    text_string = str(text_string).strip()
    prefixes = list(group_map.keys())

    # Matches prefix + optional space/dash/underscore + numeric digits
    pattern = rf"\b({'|'.join(prefixes)})\s*[-_]?\s*(\d+)"
    match = re.search(pattern, text_string, re.IGNORECASE)

    if match:
        pref = match.group(1).upper()
        num_id = match.group(2)
        group_name = group_map[pref]
        return pref, num_id, group_name

    return None, None, None


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
    best = (0, 0, 0, 0)
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


def run_batch_bending_analysis(input_folder, group_map):
    print(f"\n--- Processing Raw Data in Folder: {input_folder} ---")
    date_str = datetime.now().strftime("%m%d%y")

    plot_folder = os.path.join(input_folder, "Fz_Displacement_Analysis")
    os.makedirs(plot_folder, exist_ok=True)

    excel_filename = f"Fz_Displacement_Analysis_{date_str}.xlsx"
    txt_files = glob.glob(os.path.join(input_folder, FILE_GLOB_PATTERN))
    if not txt_files:
        return None

    results = []

    for path in txt_files:
        base_name = os.path.basename(path).replace(".txt", "")

        # Check if the file follows your naming system
        pref, num_id, g_name = parse_id_and_group(base_name, group_map)
        if not pref:
            print(
                f"Skipping file {base_name}: Doesn't match group configuration prefixes.")
            continue

        try:
            df = read_bending_txt(path)
            displacement = df["Position (z), mm"].values
            load = df["Fz, N"].values

            load_smooth = np.convolve(load, np.ones(3) / 3, mode="same")
            DISPLACEMENT_LIMIT = 1.75
            valid_range_mask = displacement <= DISPLACEMENT_LIMIT

            load_for_peak = load_smooth[valid_range_mask]
            constrained_max = np.max(load_for_peak) if len(
                load_for_peak) > 0 else np.max(load_smooth)

            candidates = np.where(
                (load_smooth >= 0.5 * constrained_max) & (valid_range_mask))[0]
            max_idx = candidates[np.argmax(load[candidates])] if len(
                candidates) > 0 else np.argmax(load_smooth)

            max_load = load[max_idx]
            post_max_load = load[max_idx:]
            NEAR_ZERO_THRESHOLD = 0.5
            drop_threshold_value = max_load * 0.80

            zero_drop_indices = np.where(
                post_max_load <= NEAR_ZERO_THRESHOLD)[0]
            if len(zero_drop_indices) > 0:
                fail_idx = max_idx + zero_drop_indices[0]
            else:
                has_dropped_indices = np.where(
                    post_max_load < drop_threshold_value)[0]
                if len(has_dropped_indices) > 0:
                    search_start = has_dropped_indices[0]
                    search_load = post_max_load[search_start:]
                    diff_post = np.diff(np.convolve(
                        search_load, np.ones(3) / 3, mode="same"))
                    local_min_indices = np.where(diff_post >= 0)[0]
                    fail_idx = max_idx + search_start + \
                        (local_min_indices[0] if len(
                            local_min_indices) > 0 else len(search_load) - 1)
                else:
                    fail_idx = len(load) - 1

            pre_max_load = load[:max_idx]
            toe_mask = pre_max_load >= (TOE_LOAD_FRACTION * max_load)
            toe_indices = np.where(toe_mask)[0]

            start_idx = toe_indices[0] if len(toe_indices) > 0 else 0
            adj_disp = displacement - displacement[start_idx]

            disp_at_failure = adj_disp[fail_idx]
            load_at_failure = load[fail_idx]

            disp_slope_candidates = (
                displacement[:max_idx][toe_mask] - displacement[start_idx])
            load_slope_candidates = pre_max_load[toe_mask]

            if len(disp_slope_candidates) < LINEAR_WINDOW_POINTS:
                stiffness, intercept, idx0, idx1 = np.nan, np.nan, 0, 0
            else:
                stiffness, intercept, idx0, idx1 = dominant_linear_region(
                    disp_slope_candidates, load_slope_candidates, window=LINEAR_WINDOW_POINTS
                )

            if hasattr(np, 'trapezoid'):
                energy = np.trapezoid(
                    load[start_idx: fail_idx + 1], adj_disp[start_idx: fail_idx + 1])
            else:
                energy = np.trapz(
                    load[start_idx: fail_idx + 1], adj_disp[start_idx: fail_idx + 1])

            plt.figure(figsize=(7, 5))
            plt.plot(adj_disp, load, color="black", label="Data")
            if (idx1 - idx0) > 5 and np.isfinite(stiffness):
                x_lin = disp_slope_candidates[idx0:idx1]
                plt.plot(x_lin, stiffness * x_lin + intercept,
                         color="red", label=f"Stiffness: {stiffness:.2f} N/mm")

            plt.scatter(adj_disp[max_idx], max_load,
                        color="green", label="Max Load")
            plt.scatter(disp_at_failure, load_at_failure,
                        color="purple", label="Failure")
            plt.axvline(x=0, color="blue", linestyle="--", label="Toe End")
            plt.fill_between(adj_disp[start_idx: fail_idx + 1],
                             load[start_idx: fail_idx + 1], alpha=0.2, color="orange")
            plt.title(base_name)
            plt.legend()
            plt.savefig(os.path.join(
                plot_folder, f"{base_name}.png"), dpi=SAVE_PNG_DPI)
            plt.close()

            results.append(
                {
                    "Standardized_ID": f"{pref}{num_id}",
                    "Max_Load_N": round(max_load, 4),
                    "Stiffness_N_per_mm": round(stiffness, 4),
                    "Energy_to_Failure_Nmm": round(energy, 4),
                    "Displacement_at_Failure_mm": round(disp_at_failure, 4),
                }
            )
        except Exception as e:
            print(f"Error processing mechanics for {base_name}: {e}")

    if not results:
        return None

    output_excel_path = os.path.join(input_folder, excel_filename)
    pd.DataFrame(results).to_excel(output_excel_path, index=False)
    return output_excel_path


# ===========================================================
# PART 2: MASTER MERGING & CONSOLIDATION
# ===========================================================

def sync_data_to_master(analysis_excel_path, master_file, measurement_file, group_map):
    if not analysis_excel_path or not os.path.exists(analysis_excel_path):
        return

    df_mach = pd.read_excel(analysis_excel_path)
    df_meas = pd.DataFrame()

    if os.path.exists(measurement_file):
        try:
            df_meas = pd.read_excel(measurement_file)
        except Exception as e:
            print(f"Could not read measurement file: {e}")

    wb = openpyxl.load_workbook(master_file, keep_links=True)
    ws = wb.active

    headers = [str(ws.cell(row=1, column=c).value).strip()
               for c in range(1, ws.max_column + 1)]

    def find_col(name_options):
        for opt in name_options:
            if opt in headers:
                return headers.index(opt) + 1
        return None

    # Identify output positions based on existing workbook structure
    col_len = find_col(["Length", "Avg. Length", "Avg_Length"]) or 2
    col_diam = find_col(["Diameter", "Avg. Diameter", "Avg_Diameter"]) or 3
    col_thick = find_col(["Thickness", "Avg. Thickness", "Avg_Thickness"]) or 4
    col_max = find_col(["Maximum Load", "Max_Load", "Max_Load_N"]) or 5
    col_stiff = find_col(["Stiffness", "Stiffness_N_per_mm"]) or 6
    col_energy = find_col(
        ["Energy to Failure", "Energy_to_Failure", "Energy_to_Failure_Nmm"]) or 7
    col_disp = find_col(["Displacement at Failure",
                        "Displacement_at_Failure", "Displacement_at_Failure_mm"]) or 8

    # Scan rows to match code patterns anywhere in column 1 or 2
    for row in range(2, ws.max_row + 1):
        cell_val1 = ws.cell(row=row, column=1).value
        cell_val2 = ws.cell(row=row, column=2).value

        pref, num_id, _ = parse_id_and_group(cell_val1, group_map)
        if not pref:
            pref, num_id, _ = parse_id_and_group(cell_val2, group_map)

        if not pref:
            continue

        lookup_id = f"{pref}{num_id}"

        # Sync geometric morphometry properties
        if not df_meas.empty:
            for idx, m_row in df_meas.iterrows():
                m_pref, m_num, _ = parse_id_and_group(m_row.iloc[0], group_map)
                if m_pref and f"{m_pref}{m_num}" == lookup_id:
                    ws.cell(
                        row=row, column=col_len).value = m_row.iloc[1] if df_meas.shape[1] > 1 else None
                    ws.cell(
                        row=row, column=col_diam).value = m_row.iloc[8] if df_meas.shape[1] > 8 else None
                    ws.cell(
                        row=row, column=col_thick).value = m_row.iloc[12] if df_meas.shape[1] > 12 else None
                    break

        # Sync automated testing metrics
        mach_row = df_mach[df_mach["Standardized_ID"] == lookup_id]
        if not mach_row.empty:
            ws.cell(
                row=row, column=col_max).value = mach_row.iloc[0]["Max_Load_N"]
            ws.cell(
                row=row, column=col_stiff).value = mach_row.iloc[0]["Stiffness_N_per_mm"]
            ws.cell(
                row=row, column=col_energy).value = mach_row.iloc[0]["Energy_to_Failure_Nmm"]
            ws.cell(
                row=row, column=col_disp).value = mach_row.iloc[0]["Displacement_at_Failure_mm"]

    wb.save(master_file)


# ===========================================================
# PART 3: SEGREGATED SUB-STUDY CSV GENERATION (RE-PIVOTING)
# ===========================================================

def generate_segregated_csvs(master_path, export_dir, group_map):
    """
    Scans the updated Master sheet, pulls valid group codes dynamically,
    and reshapes the data into side-by-side tables named strictly by metric.
    """
    if master_path.endswith('.csv'):
        df = pd.read_csv(master_path)
    else:
        df = pd.read_excel(master_path)

    headers = [str(c).strip() for c in df.columns]

    def find_col_name(name_options):
        for opt in name_options:
            if opt in headers:
                return opt
            for h in headers:
                if opt.lower() in h.lower():
                    return h
        return None

    # Clean, standardized keys that will serve as your exact filenames
    metric_mapping = {
        "Avg_Length": find_col_name(["Length", "Avg_Length", "Avg. Length"]),
        "Avg_Diameter": find_col_name(["Diameter", "Avg_Diameter", "Avg. Diameter"]),
        "Avg_Thickness": find_col_name(["Thickness", "Avg_Thickness", "Avg. Thickness"]),
        "Max_Load": find_col_name(["Max_Load", "Maximum Load", "Max Load", "Max_Load_N"]),
        "Stiffness": find_col_name(["Stiffness", "Stiffness_N_per_mm", "Stiffness (N/mm)"]),
        "Energy_to_Failure": find_col_name(["Energy", "Energy_to_Failure", "Energy to Failure"]),
        "Displacement_at_Failure": find_col_name(["Displacement", "Displacement_at_Failure"])
    }

    parsed_rows = []

    for idx, row in df.iterrows():
        matched_pref, matched_num, matched_gname = None, None, None

        for col_idx in range(min(3, len(df.columns))):
            p, n, g = parse_id_and_group(row.iloc[col_idx], group_map)
            if p:
                matched_pref, matched_num, matched_gname = p, n, g
                break

        if not matched_pref:
            continue

        row_data = {
            "ID_Num": int(matched_num),
            "Group": matched_gname
        }

        for metric_key, actual_col in metric_mapping.items():
            row_data[metric_key] = row[actual_col] if actual_col else np.nan

        parsed_rows.append(row_data)

    if not parsed_rows:
        print("Notice: No matching group IDs could be extracted from master rows.")
        return

    df_parsed = pd.DataFrame(parsed_rows)
    expected_columns = list(group_map.values())
    os.makedirs(export_dir, exist_ok=True)

    # Re-pivot each found metric column into side-by-side dataframes
    for metric_name in metric_mapping.keys():
        if df_parsed[metric_name].isna().all():
            continue

        pivot_df = df_parsed.pivot_table(
            index="ID_Num",
            columns="Group",
            values=metric_name,
            aggfunc="first"
        )

        pivot_df = pivot_df.reindex(columns=expected_columns)
        pivot_df = pivot_df.sort_index()

        # Cleans up file naming so it outputs exactly as "Max_Load.csv", "Stiffness.csv", etc.
        out_csv_path = os.path.join(export_dir, f"{metric_name}.csv")

        with open(out_csv_path, 'w', encoding='utf-8') as f:
            f.write("sep=,\n")

        pivot_df.to_csv(out_csv_path, mode='a', sep=",", index=True)
        print(f"Generated side-by-side table: {out_csv_path}")


# ===========================================================
# PART 4: PIPELINE EXECUTION ENGINE
# ===========================================================

def execute_single_study_pipeline(data_folder, master_path, measurement_path, csv_out_dir, group_map):
    root_path = Path(data_folder)

    # 1. Complete mechanical force crunching
    analysis_excel = run_batch_bending_analysis(str(root_path), group_map)
    if not analysis_excel:
        print("Execution halted: No valid data files were evaluated.")
        return False

    # 2. Append values straight into the workbook
    if os.path.exists(master_path):
        sync_data_to_master(analysis_excel, master_path,
                            measurement_path, group_map)

        # 3. Drop a copy of the master spreadsheet into the master directory
        try:
            df_final = pd.read_excel(master_path)
            master_dir = os.path.dirname(master_path)

            # CHANGED: The file is now explicitly named with your study prefix
            compiled_filename = "IFS+SHP099+Medigel_LFemurMaster_Compiled.csv"
            compiled_path = os.path.join(master_dir, compiled_filename)

            df_final.to_csv(compiled_path, index=False, sep=",")
            print(f"Main master file compiled and saved to: {compiled_path}")
        except Exception as e:
            print(f"Failed to generate compiled master copy: {e}")

        # 4. Generate individual pivoted matrices
        try:
            generate_segregated_csvs(master_path, csv_out_dir, group_map)
        except Exception as e:
            print(f"Error compiling side-by-side metrics: {e}")

    return True


# ===========================================================
# PART 5: USER INTERFACE SCREEN DESIGN
# ===========================================================

def launch_interface():
    root = tk.Tk()
    root.title("Single Study Bending Pipeline")
    root.geometry("680x400")
    root.resizable(False, False)

    path_raw = tk.StringVar(value=STUDY_CONFIG["raw_data_root"])
    path_master = tk.StringVar(value=STUDY_CONFIG["master_file"])
    path_meas = tk.StringVar(value=STUDY_CONFIG["measurement_file"])
    path_out = tk.StringVar(value=STUDY_CONFIG["output_folder"])

    def browse_raw():
        f = filedialog.askdirectory(title="Select Raw Text Directory")
        if f:
            path_raw.set(f)

    def browse_master():
        f = filedialog.askopenfilename(
            title="Select Target Master Workbook", filetypes=[("Excel Files", "*.xlsx")])
        if f:
            path_master.set(f)

    def browse_meas():
        f = filedialog.askopenfilename(title="Select Dimensions Ledger", filetypes=[
                                       ("Excel Files", "*.xlsx")])
        if f:
            path_meas.set(f)

    def browse_out():
        f = filedialog.askdirectory(
            title="Select Destination Folder for the 7 Metric CSVs")
        if f:
            path_out.set(f)

    def run_pipeline():
        root.title("Processing Single Study Matrix Data... Please wait.")
        root.update()

        current_raw = path_raw.get()
        current_master = path_master.get()
        current_meas = path_meas.get()
        current_out = path_out.get()

        # Save values to config JSON so changes are remembered
        save_study_config(current_raw, current_master,
                          current_meas, current_out)

        try:
            success = execute_single_study_pipeline(
                data_folder=current_raw,
                master_path=current_master,
                measurement_path=current_meas,
                csv_out_dir=current_out,
                group_map=STUDY_CONFIG["group_map"]
            )
            if success:
                messagebox.showinfo(
                    "Success", "Calculations completed. Master updated and 7 metric CSVs exported successfully.")
            else:
                messagebox.showwarning(
                    "Incomplete", "No analysis calculations could be finalized.")
        except Exception as e:
            messagebox.showerror(
                "Execution Error", f"An unhandled error occurred:\n{e}")
        finally:
            root.title("Single Study Bending Pipeline")

    tk.Label(root, text="Single Study Mechanical Analysis",
             font=("Arial", 13, "bold")).pack(pady=15)

    fields_frame = tk.Frame(root)
    fields_frame.pack(fill="both", expand=True, padx=25)

    def add_ui_row(label_text, variable, command):
        row = tk.Frame(fields_frame)
        row.pack(fill="x", pady=6)
        tk.Label(row, text=label_text, width=22, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=variable,
                 width=50).pack(side="left", padx=5)
        tk.Button(row, text="Browse...", command=command).pack(side="left")

    add_ui_row("Raw Data Folder:", path_raw, browse_raw)
    add_ui_row("Master Excel File:", path_master, browse_master)
    add_ui_row("Physical Measurements Excel File:", path_meas, browse_meas)
    add_ui_row("CSV Export Folder:", path_out, browse_out)

    tk.Button(root, text="Execute Workflow", bg="#2E7D32", fg="white", font=(
        "Arial", 11, "bold"), padx=30, pady=8, command=run_pipeline).pack(pady=15)

    root.mainloop()


if __name__ == "__main__":
    launch_interface()
