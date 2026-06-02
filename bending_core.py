"""Shared bending analysis primitives — imported by all analysis scripts."""
from io import StringIO
import numpy as np
import pandas as pd

TOE_LOAD_FRACTION       = 0.05
LINEAR_WINDOW_POINTS    = 90
MIN_R2                  = 0.995
DISPLACEMENT_LIMIT      = 1.75
NEAR_ZERO_THRESHOLD     = 0.5
DROP_THRESHOLD_FRACTION = 0.80
PEAK_CANDIDATE_FRACTION = 0.50


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
        except (ValueError, IndexError):
            header = parts
            data_lines = data_lines[i + 1:]
            break

    df = pd.read_csv(StringIO("".join(data_lines)), sep="\t", names=header, engine="python")
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["Fz, N", "Position (z), mm"], how="any")
    if "Fx" in df.columns:
        df["Fz, N"] = df["Fx"]
    if "Position (z)" in df.columns:
        df["Position (z), mm"] = df["Position (z)"]
    df["Fz, N"] = -df["Fz, N"]
    return df


def dominant_linear_region(x, y, window=LINEAR_WINDOW_POINTS, min_r2=MIN_R2):
    best_len = 0
    best = (0, 0, 0, 0)
    n = len(x)
    if n < window:
        return 0, 0, 0, n
    for start in range(0, n - window, 2):
        end = start + window
        m, b = np.polyfit(x[start:end], y[start:end], 1)
        r2 = np.corrcoef(x[start:end], y[start:end])[0, 1] ** 2
        if r2 >= min_r2:
            while end < n:
                next_end = min(end + 5, n)
                m_ext, b_ext = np.polyfit(x[start:next_end], y[start:next_end], 1)
                if np.corrcoef(x[start:next_end], y[start:next_end])[0, 1] ** 2 < min_r2:
                    break
                m, b, end = m_ext, b_ext, next_end
            if (end - start) > best_len:
                best_len = end - start
                best = (m, b, start, end)
    return best
