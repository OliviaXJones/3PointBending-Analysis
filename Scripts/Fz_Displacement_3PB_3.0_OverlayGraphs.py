import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import StringIO
from datetime import datetime

# ================= FUNCTIONS (MATCHING YOUR BATCH ANALYZER) =================
def read_bending_txt(filepath):
    with open(filepath, 'r', errors='ignore') as f:
        lines = f.readlines()
    data_start = next(i + 1 for i, line in enumerate(lines) if "<DATA>" in line)
    data_end = next(i for i, line in enumerate(lines) if "<END DATA>" in line)
    data_lines = lines[data_start:data_end]
    header = None
    for i, line in enumerate(data_lines):
        try:
            float(line.strip().split('\t')[0])
        except:
            header = line.strip().split('\t')
            data_lines = data_lines[i+1:]
            break
    df = pd.read_csv(StringIO("".join(data_lines)), sep="\t", names=header, engine='python')
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna(subset=['Fz, N', 'Position (z), mm'], how='any')
    if 'Fx' in df.columns: df['Fz, N'] = df['Fx']
    if 'Position (z)' in df.columns: df['Position (z), mm'] = df['Position (z)']
    df['Fz, N'] = -df['Fz, N']
    return df

def dominant_linear_region(x, y, window=90, min_r2=0.995):
    best_len = 0
    best = (0, 0, 0, 0)
    n = len(x)
    if n < window: return 0, 0, 0, n
    for start in range(0, n - window, 2): 
        end = start + window
        x_seg, y_seg = x[start:end], y[start:end]
        m, b = np.polyfit(x_seg, y_seg, 1)
        r2 = np.corrcoef(x_seg, y_seg)[0, 1]**2
        if r2 >= min_r2:
            while end < n:
                next_end = min(end + 5, n)
                m_ext, b_ext = np.polyfit(x[start:next_end], y[start:next_end], 1)
                if (np.corrcoef(x[start:next_end], y[start:next_end])[0, 1]**2) < min_r2: break
                m, b, end = m_ext, b_ext, next_end
            if (end - start) > best_len:
                best_len = end - start
                best = (m, b, start, end)
    return best

# ================= DYNAMIC INPUTS =================
print("--- Advanced Comparison Plotter (Sync Logic) ---")
custom_title = input("Enter the Graph Title: ").strip()

save_dir = input("Paste the FOLDER path where you want to save the graph: ").strip().replace('"', '')
if not os.path.exists(save_dir):
    os.makedirs(save_dir)

num_files = int(input("How many files do you want to overlay? "))

files_data = []
for i in range(num_files):
    print(f"\n--- Entry {i+1} ---")
    path = input(f"Paste path for File {i+1}: ").strip().replace('"', '')
    color = input(f"Hex code for File {i+1}: ").strip()
    files_data.append((path, color))

# ================= PLOTTING =================
plt.figure(figsize=(15, 10))
date_str = datetime.now().strftime("%m%d%y")
fail_points = []

# Global font size settings
plt.rcParams.update({'font.size': 13})

for path, color in files_data:
    if not os.path.exists(path):
        print(f"❌ File not found: {path}")
        continue
        
    try:
        df = read_bending_txt(path)
        disp = df['Position (z), mm'].values
        load = df['Fz, N'].values
        fname = os.path.basename(path).replace("_", "\\_")

        # 1. Peak Detection (Matching your constraints)
        load_smooth = np.convolve(load, np.ones(3)/3, mode='same')
        DISPLACEMENT_LIMIT = 1.75
        valid_range_mask = disp <= DISPLACEMENT_LIMIT
        load_for_peak = load_smooth[valid_range_mask]
        constrained_max = np.max(load_for_peak) if len(load_for_peak) > 0 else np.max(load_smooth)
        candidates = np.where((load_smooth >= 0.5 * constrained_max) & (valid_range_mask))[0]
        max_idx = candidates[np.argmax(load[candidates])] if len(candidates) > 0 else np.argmax(load_smooth)
        max_load, disp_at_max = load[max_idx], disp[max_idx]

        # 2. Robust Failure Detection (Synchronized with your analyzer)
        post_max_load = load[max_idx:]
        NEAR_ZERO_THRESHOLD = 0.5 
        drop_threshold_value = max_load * 0.80 # 20% drop

        zero_drop_indices = np.where(post_max_load <= NEAR_ZERO_THRESHOLD)[0]
        if len(zero_drop_indices) > 0:
            fail_idx = max_idx + zero_drop_indices[0]
        else:
            has_dropped_indices = np.where(post_max_load < drop_threshold_value)[0]
            if len(has_dropped_indices) > 0:
                search_start = has_dropped_indices[0]
                search_load = post_max_load[search_start:]
                diff_post = np.diff(np.convolve(search_load, np.ones(3)/3, mode='same'))
                local_min_indices = np.where(diff_post >= 0)[0]
                fail_idx = max_idx + search_start + (local_min_indices[0] if len(local_min_indices) > 0 else len(search_load)-1)
            else:
                fail_idx = len(load) - 1

        fail_points.append((disp[fail_idx], load[fail_idx]))

        # 3. Stiffness
        pre_max_disp, pre_max_load = disp[:max_idx], load[:max_idx]
        toe_mask = pre_max_load >= (0.05 * max_load) # TOE_LOAD_FRACTION = 0.05
        stiffness, intercept, idx0, idx1 = dominant_linear_region(
            pre_max_disp[toe_mask], pre_max_load[toe_mask], window=90
        )

        # 4. Legend label
        legend_label = r"$\mathbf{" + fname + r"}$" + f"\nSlope = {stiffness:.2f} N/mm\nMax Force = {max_load:.2f} N"
        plt.plot(disp, load, color=color, label=legend_label, linewidth=3, alpha=0.75)
        
        # Max Load Marker
        plt.scatter(disp_at_max, max_load, color=color, edgecolor='black', s=120, zorder=5)
        
        # Slope line
        if stiffness > 0:
            slope_x = pre_max_disp[toe_mask][idx0:idx1]
            plt.plot(slope_x, stiffness * slope_x + intercept, color='black', linestyle='--', linewidth=2.5, zorder=4)

    except Exception as e:
        print(f"❌ Error processing {path}: {e}")

# Failure Points and Legend Spacer
plt.plot([], [], ' ', label="") 
for i, (f_disp, f_load) in enumerate(fail_points):
    plt.scatter(f_disp, f_load, color='red', marker='x', s=180, linewidth=4, zorder=6, label="Failure Point" if i == 0 else "")

# Styling
plt.title(custom_title, fontsize=24, fontweight='bold', pad=25)
plt.xlabel("Position (z), mm", fontsize=16, fontweight='bold')
plt.ylabel("Fz, N", fontsize=16, fontweight='bold')
plt.legend(loc='upper right', frameon=True, shadow=True, borderpad=1.2, fontsize=13, labelspacing=1.8)
plt.grid(True, linestyle=':', alpha=0.4)
plt.tick_params(axis='both', labelsize=14)

# Save
output_path = os.path.join(save_dir, f"Overlay_Comparison_{date_str}.png")
plt.savefig(output_path, dpi=300, bbox_inches='tight')
print(f"\n✅ Success! Graph saved to: {output_path}")
plt.show()