# ===========================================================
# Batch Bending Data Analyzer (Load–Displacement)
# ===========================================================

import os
import glob
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import StringIO
from datetime import datetime
from tkinter import Tk, filedialog, messagebox

root = Tk()
root.withdraw()
root.attributes('-topmost', True)

# 1. Show the instruction popup
"""messagebox.showinfo("Folder Selection",
                    "Please navigate to the folder containing 3-Point Bending Force-Displacement .txt files")"""

# 2. Open the folder selection dialog
INPUT_FOLDER = filedialog.askdirectory(
    title="Select Force-Displacement Folder")

# Cleanup the root window
root.destroy()

# Proceed only if a folder was selected
if not INPUT_FOLDER:
    print("No folder selected. Operation cancelled.")
    exit()
# Get current date in MMDDYY format
date_str = datetime.now().strftime("%m%d%y")

PLOT_FOLDER = os.path.join(INPUT_FOLDER, "Fz_Displacement_Analysis")
os.makedirs(PLOT_FOLDER, exist_ok=True)

FILE_GLOB_PATTERN = "*.txt"
SAVE_PNG_DPI = 100
TRY_EXPORT_EXCEL = True

# Dynamic Filename: Fz_Displacement_Analysis_MMDDYY.xlsx
EXCEL_FILENAME = f"Fz_Displacement_Analysis_{date_str}.xlsx"

TOE_LOAD_FRACTION = 0.05
LINEAR_WINDOW_POINTS = 90  # Points to start the search
MIN_R2 = 0.995

# ================= FUNCTIONS =================


def read_bending_txt(filepath):
    with open(filepath, 'r', errors='ignore') as f:
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
        parts = line.strip().split('\t')
        try:
            float(parts[0])
        except:
            header = parts
            data_lines = data_lines[i+1:]
            break

    df = pd.read_csv(StringIO("".join(data_lines)),
                     sep="\t", names=header, engine='python')
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df = df.dropna(subset=['Fz, N', 'Position (z), mm'], how='any')
    if 'Fx' in df.columns:
        df['Fz, N'] = df['Fx']
    if 'Position (z)' in df.columns:
        df['Position (z), mm'] = df['Position (z)']
    df['Fz, N'] = -df['Fz, N']
    return df


def dominant_linear_region(x, y, window=30, min_r2=0.995):
    """Optimized linear region search to prevent hanging."""
    best_len = 0
    best = (0, 0, 0, 0)  # m, b, start, end
    n = len(x)

    if n < window:
        return 0, 0, 0, n

    # Step through the data in small increments to speed up search
    for start in range(0, n - window, 2):
        end = start + window
        x_seg = x[start:end]
        y_seg = y[start:end]

        m, b = np.polyfit(x_seg, y_seg, 1)
        r = np.corrcoef(x_seg, y_seg)[0, 1]
        r2 = r**2

        if r2 >= min_r2:
            # Extend forward in larger chunks to save time
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


# ================= MAIN SCRIPT =================
txt_files = glob.glob(os.path.join(INPUT_FOLDER, FILE_GLOB_PATTERN))
if not txt_files:
    print(f"No TXT files found in {INPUT_FOLDER}")
    exit()

results = []
combined_fig, combined_ax = plt.subplots(figsize=(10, 7))

for path in txt_files:
    try:
        df = read_bending_txt(path)
        displacement = df['Position (z), mm'].values
        load = df['Fz, N'].values

        # Smoothing for peak detection
        load_smooth = np.convolve(load, np.ones(3)/3, mode='same')

        # Limit search to realistic bone displacement
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
        disp_at_max = displacement[max_idx]

        # 5. Robust Failure detection
        post_max_load = load[max_idx:]
        NEAR_ZERO_THRESHOLD = 0.5
        drop_threshold_value = max_load * 0.80  # 20% drop

        zero_drop_indices = np.where(post_max_load <= NEAR_ZERO_THRESHOLD)[0]
        if len(zero_drop_indices) > 0:
            fail_idx = max_idx + zero_drop_indices[0]
        else:
            has_dropped_indices = np.where(
                post_max_load < drop_threshold_value)[0]
            if len(has_dropped_indices) > 0:
                search_start = has_dropped_indices[0]
                search_load = post_max_load[search_start:]
                diff_post = np.diff(np.convolve(
                    search_load, np.ones(3)/3, mode='same'))
                local_min_indices = np.where(diff_post >= 0)[0]
                fail_idx = max_idx + search_start + \
                    (local_min_indices[0] if len(
                        local_min_indices) > 0 else len(search_load)-1)
            else:
                fail_idx = len(load) - 1

        # --- Toe filtering & Slope Locking ---
        pre_max_disp = displacement[:max_idx]
        pre_max_load = load[:max_idx]

        # Define toe mask FIRST
        toe_mask = pre_max_load >= (TOE_LOAD_FRACTION * max_load)

        toe_indices = np.where(toe_mask)[0]

        if len(toe_indices) > 0:
            start_idx = toe_indices[0]
        else:
            start_idx = 0  # fallback

        # Now define adjusted displacement
        adj_disp = displacement - displacement[start_idx]

        # Shifted reference points (FOR PLOTTING)
        adj_max_disp = adj_disp[max_idx]
        adj_fail_disp = adj_disp[fail_idx]

        # Values for reporting
        disp_at_failure = adj_fail_disp
        load_at_failure = load[fail_idx]

        # Convert candidates into SAME coordinate system as plot
        disp_slope_candidates = (
            pre_max_disp[toe_mask] - displacement[start_idx])
        load_slope_candidates = pre_max_load[toe_mask]

        # 🔒 SAFETY CHECK (ADD HERE)
        if len(disp_slope_candidates) < LINEAR_WINDOW_POINTS:
            print(
                f"Skipping stiffness fit (too few points) in {os.path.basename(path)}")
            stiffness = np.nan
            intercept = np.nan
            idx0, idx1 = 0, 0
        else:
            stiffness, intercept, idx0, idx1 = dominant_linear_region(
                disp_slope_candidates,
                load_slope_candidates,
                window=LINEAR_WINDOW_POINTS
            )

        energy = np.trapezoid(
            load[start_idx:fail_idx+1],
            adj_disp[start_idx:fail_idx+1]
        )
        # Plotting
        plt.figure(figsize=(7, 5))
        plt.plot(adj_disp, load, color='black', label='Data')
        if (idx1 - idx0) > 5 and np.isfinite(stiffness):
            x_lin = disp_slope_candidates[idx0:idx1]

            plt.plot(
                x_lin,
                stiffness * x_lin + intercept,
                color='red',
                label=f'Stiffness: {stiffness:.2f} N/mm'
            )

        plt.scatter(adj_max_disp, max_load, color='green', label='Max Load')
        plt.scatter(adj_fail_disp, load_at_failure,
                    color='purple', label='Failure')
        plt.axvline(x=0, color='blue', linestyle='--', label='Toe End (Start)')
        plt.fill_between(
            adj_disp[start_idx:fail_idx+1],
            load[start_idx:fail_idx+1],
            alpha=0.2,
            color='orange'
        )
        plt.title(os.path.basename(path))
        plt.legend()
        plt.savefig(os.path.join(PLOT_FOLDER, os.path.basename(
            path).replace(".txt", ".png")), dpi=SAVE_PNG_DPI)
        plt.close()

        results.append({
            "Filename": os.path.basename(path),
            "Max_Load_N": round(max_load, 4),
            "Stiffness_N_per_mm": round(stiffness, 4),
            "Energy_to_Failure_Nmm": round(energy, 4),
            "Displacement_at_Failure_mm": round(disp_at_failure, 4)
        })
        print(f"Processed: {os.path.basename(path)}")

    except Exception as e:
        print(f"Error {os.path.basename(path)}: {e}")

# Save results
pd.DataFrame(results).to_excel(os.path.join(
    INPUT_FOLDER, EXCEL_FILENAME), index=False)
