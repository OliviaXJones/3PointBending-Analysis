"""
Convert workflow CSV output folders into GraphPad Prism .pzfx files.

Each CSV (one metric × sex × age) becomes one Prism table.
Groups (Wildtype, Mutant, etc.) map to Prism columns.
Mouse IDs (ID_Num) map to Prism row titles.
"""

import os
import glob
import pandas as pd
from pzfx import write_pzfx

_EXPORT_FOLDERS = [
    "Femur_Analysis_By_Genotype",
    "Femur_Analysis_By_Lineage",
    "Tibia_Analysis_By_Genotype",
    "Tibia_Analysis_By_Lineage",
    "Humerus_Analysis_By_Genotype",
    "Anteroposterior_Femur",
    "Mediolateral_Femur",
    "Proximal_Tibia",
    "Distal_Tibia",
]


def _load_csv_tables(csv_folder, n_digits=4):
    """Load all CSVs in a folder into a dict of {table_name: DataFrame}."""
    tables = {}
    for path in sorted(glob.glob(os.path.join(csv_folder, "*.csv"))):
        table_name = os.path.splitext(os.path.basename(path))[0]
        df = pd.read_csv(path, index_col=0)
        df = df.dropna(how="all")
        df = df.loc[:, df.notna().any()]
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        if df.empty or df.columns.empty:
            continue
        tables[table_name] = df
    return tables


def csvs_to_pzfx(csv_folder, output_path, n_digits=4):
    """
    Load every .csv in csv_folder and write them as tables in a single .pzfx file.
    Returns [output_path] if written, [] if the folder was empty.
    """
    tables = _load_csv_tables(csv_folder, n_digits)
    if not tables:
        return []
    write_pzfx(tables, output_path, row_names=True, n_digits=n_digits)
    return [output_path]


def workflow_output_to_pzfx(csv_out_dir, study_name):
    """
    Scan csv_out_dir for any *_Analysis_By_Genotype subfolders and convert
    each to a .pzfx file saved alongside the CSVs.

    Returns a list of paths to the created .pzfx files.
    """
    created = []

    for sex in ["Male", "Female"]:
        sex_dir = os.path.join(csv_out_dir, sex)
        if not os.path.isdir(sex_dir):
            continue
        for folder_name in _EXPORT_FOLDERS:
            folder = os.path.join(sex_dir, folder_name)
            if not os.path.isdir(folder):
                continue
            out_path = os.path.join(csv_out_dir, f"{study_name}_{folder_name}_{sex}.pzfx")
            try:
                written = csvs_to_pzfx(folder, out_path)
                for p in written:
                    created.append(p)
                    print(f"Prism export: {p}")
            except Exception as e:
                print(f"Prism export failed for {sex}/{folder_name}: {e}")

    return created
