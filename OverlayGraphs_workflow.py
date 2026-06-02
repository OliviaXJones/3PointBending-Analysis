"""Overlay graph generation — called from Main_Dashboard, streamlit_app, or the standalone CLI."""
import os
from datetime import datetime

DEFAULT_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]

import matplotlib.pyplot as plt
import numpy as np

from bending_core import (
    read_bending_txt, dominant_linear_region,
    TOE_LOAD_FRACTION, LINEAR_WINDOW_POINTS,
    DISPLACEMENT_LIMIT, NEAR_ZERO_THRESHOLD, DROP_THRESHOLD_FRACTION,
)


def run_overlay(files_data, title, save_dir, dpi=300):
    """Generate a toe-adjusted overlay comparison graph.

    files_data : list of (path, label, color_hex)
    Returns    : output_path (str)
    """
    os.makedirs(save_dir, exist_ok=True)

    plt.rcParams.update({"font.size": 13})
    fig, ax = plt.subplots(figsize=(15, 10))
    fail_points = []

    for path, label, color in files_data:
        if not os.path.exists(path):
            print(f"File not found: {path}")
            continue
        try:
            df = read_bending_txt(path)
            displacement = df["Position (z), mm"].values
            load = df["Fz, N"].values

            # Peak detection (identical to batch analyzer)
            load_smooth = np.convolve(load, np.ones(3) / 3, mode="same")
            valid_mask = displacement <= DISPLACEMENT_LIMIT
            load_for_peak = load_smooth[valid_mask]
            constrained_max = np.max(load_for_peak) if len(load_for_peak) > 0 else np.max(load_smooth)
            candidates = np.where((load_smooth >= 0.5 * constrained_max) & valid_mask)[0]
            max_idx = candidates[np.argmax(load[candidates])] if len(candidates) > 0 else np.argmax(load_smooth)
            max_load = load[max_idx]

            # Failure detection (identical to batch analyzer)
            post_max_load = load[max_idx:]
            zero_drop = np.where(post_max_load <= NEAR_ZERO_THRESHOLD)[0]
            if len(zero_drop) > 0:
                fail_idx = max_idx + zero_drop[0]
            else:
                dropped = np.where(post_max_load < max_load * DROP_THRESHOLD_FRACTION)[0]
                if len(dropped) > 0:
                    s = dropped[0]
                    sl = post_max_load[s:]
                    diff = np.diff(np.convolve(sl, np.ones(3) / 3, mode="same"))
                    lm = np.where(diff >= 0)[0]
                    fail_idx = max_idx + s + (lm[0] if len(lm) > 0 else len(sl) - 1)
                else:
                    fail_idx = len(load) - 1

            # Toe-adjusted x-axis — aligns all curves at x=0 (toe end), same as batch analyzer
            pre_max_load = load[:max_idx]
            toe_mask = pre_max_load >= (TOE_LOAD_FRACTION * max_load)
            toe_indices = np.where(toe_mask)[0]
            start_idx = toe_indices[0] if len(toe_indices) > 0 else 0
            adj_disp = displacement - displacement[start_idx]

            # Stiffness
            disp_slope = adj_disp[:max_idx][toe_mask]
            load_slope = pre_max_load[toe_mask]
            if len(disp_slope) >= LINEAR_WINDOW_POINTS:
                stiffness, intercept, idx0, idx1 = dominant_linear_region(
                    disp_slope, load_slope, window=LINEAR_WINDOW_POINTS
                )
            else:
                stiffness, intercept, idx0, idx1 = 0, 0, 0, 0

            legend_text = f"{label}\nSlope = {stiffness:.2f} N/mm  |  Max = {max_load:.2f} N"
            ax.plot(adj_disp, load, color=color, label=legend_text, linewidth=2.5, alpha=0.8)
            ax.scatter(adj_disp[max_idx], max_load, color=color, edgecolor="black", s=120, zorder=5)

            if stiffness > 0 and (idx1 - idx0) > 5:
                slope_x = disp_slope[idx0:idx1]
                ax.plot(slope_x, stiffness * slope_x + intercept,
                        color="black", linestyle="--", linewidth=2, zorder=4)

            fail_points.append((adj_disp[fail_idx], load[fail_idx]))

        except Exception as e:
            print(f"Error processing {path}: {e}")

    ax.plot([], [], " ", label="")
    for i, (f_disp, f_load) in enumerate(fail_points):
        ax.scatter(f_disp, f_load, color="red", marker="x", s=180,
                   linewidth=4, zorder=6, label="Failure Point" if i == 0 else "")

    ax.set_title(title, fontsize=24, fontweight="bold", pad=25)
    ax.set_xlabel("Position (z), mm", fontsize=16, fontweight="bold")
    ax.set_ylabel("Fz, N", fontsize=16, fontweight="bold")
    ax.legend(loc="upper right", frameon=True, shadow=True,
              borderpad=1.2, fontsize=12, labelspacing=1.5)
    ax.grid(True, linestyle=":", alpha=0.4)
    ax.tick_params(axis="both", labelsize=13)

    date_str = datetime.now().strftime("%m%d%y")
    output_path = os.path.join(save_dir, f"Overlay_{date_str}.png")
    fig.savefig(output_path, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    return output_path
