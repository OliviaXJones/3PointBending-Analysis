"""
FKBP5 New study pipeline.

Mouse code format:  [GenotypeCode].[Age].[SexID]
  W = Wildtype      e.g.  W.12.F23  → Wildtype, 12-week, Female #23
  Z = Heterozygous  e.g.  Z.4.M2   → Heterozygous, 4-week, Male #2
  X = Mutant        e.g.  X.6.F1   → Mutant, 6-week, Female #1

Differences from FKBP5_workflow:
  - Genotype letter codes W / Z / X (same dot-separated format, parsed identically)
  - No lineage folder — genotype-only CSV output
  - Separate default paths and compiled-CSV filename prefix
"""

import os
from datetime import datetime
from pathlib import Path

import pandas as pd

# Re-use all shared infrastructure from the original FKBP5 workflow
from FKBP5Heat_workflow import (
    run_batch_bending_analysis,
    sync_data_to_master,
    parse_mouse_code,
    parse_anatomical_diameters,
    determine_bone_type,
    FILE_GLOB_PATTERN,
)

# ===========================================================
# DEFAULT CONFIGURATION
# ===========================================================
DEFAULT_RAW_DATA_ROOT    = r"F:\3-Point Bending\FKBP5 New 2026\FKBP5New_Tibia_MMDDYY"
DEFAULT_TIBIA_MASTER_FILE = r"F:\3-Point Bending\FKBP5 New 2026\FKBP5New_3-PointBendingTibiaMaster.xlsx"
DEFAULT_FEMUR_MASTER_FILE = r"F:\3-Point Bending\FKBP5 New 2026\FKBP5New_3-PointBendingFemurMaster.xlsx"
DEFAULT_MEASUREMENT_FILE  = r"F:\3-Point Bending\FKBP5 New 2026\FKBP5New_Tibia+Femur_MMDDYY.xlsx"
DEFAULT_CSV_OUTPUT_DIR    = r"F:\3-Point Bending\FKBP5 New 2026\FKBP5New_CSVFiles"

# Genotype code → display name (used for column ordering in output CSVs)
GENOTYPE_MAP  = {"W": "Wildtype", "Z": "Heterozygous", "X": "Mutant"}
SORT_PRIORITY = ["Wildtype", "Heterozygous", "Mutant"]


# ===========================================================
# GENOTYPE-ONLY TABLE GENERATION (no lineage folder)
# ===========================================================

def generate_grouped_tables(df, base_dir, bone_target):
    metrics = [
        f"Avg. {bone_target} Length",
        f"Avg. {bone_target} Diameter",
        f"Avg. {bone_target} Thickness",
        "Maximum Load", "Stiffness", "Energy to Failure", "Displacement at Failure",
    ]
    fallback_metrics = [
        "Avg. Tibia Length", "Avg. Tibia Diameter", "Avg. Tibia Thickness",
        "Avg. Femur Length", "Avg. Femur Diameter", "Avg. Femur Thickness",
        "Maximum Load", "Stiffness", "Energy to Failure", "Displacement at Failure",
    ]

    if "Age_Extracted" not in df.columns:
        return

    unique_ages = df["Age_Extracted"].dropna().unique()

    active_metrics = [m for m in metrics if m in df.columns] or \
                     [m for m in fallback_metrics if m in df.columns]

    for metric in active_metrics:
        df[metric] = pd.to_numeric(df[metric], errors="coerce")

    for metric in active_metrics:
        for age in unique_ages:
            for sex in ["Male", "Female"]:
                subset = df[
                    (df["Sex_Extracted"] == sex) &
                    (df["Age_Extracted"] == age) &
                    (df["Progeny_Group"].isin(SORT_PRIORITY))
                ].copy()
                if subset.empty:
                    continue

                table = subset.pivot_table(
                    index="ID_Num", columns="Progeny_Group", values=metric)
                existing_cols = [c for c in SORT_PRIORITY if c in table.columns]
                table = table.reindex(columns=existing_cols)
                clean_metric = metric.replace("_", " ").replace(".", "")
                sex_geno_dir = os.path.join(base_dir, sex, f"{bone_target}_Analysis_By_Genotype")
                os.makedirs(sex_geno_dir, exist_ok=True)
                table.to_csv(os.path.join(sex_geno_dir, f"{age}Wks {clean_metric}.csv"))


# ===========================================================
# SHEET COMPILATION (identical logic, calls local generate fn)
# ===========================================================

def process_all_sheets(master_path, structure_type, csv_out_dir, bone_target):
    excel_data = pd.ExcelFile(master_path)
    all_compiled_data = []

    headers = [
        "Mouse Code",
        f"Avg. {bone_target} Length",
        f"Avg. {bone_target} Diameter",
        f"Avg. {bone_target} Thickness",
        "Maximum Load", "Stiffness", "Energy to Failure", "Displacement at Failure",
    ]

    for sheet in excel_data.sheet_names:
        if sheet.lower() in ("summary", "notes", "calculations"):
            continue

        if structure_type == "Split Genders Male/Female":
            raw_df = pd.read_excel(master_path, sheet_name=sheet, header=None, skiprows=1)
            if raw_df.shape[1] < 17:
                part = raw_df.iloc[:, 0:8].copy()
                part.columns = headers[:part.shape[1]]
                sheet_combined = part
            else:
                males   = raw_df.iloc[:, 0:8].copy()
                females = raw_df.iloc[:, 9:17].copy()
                males.columns = females.columns = headers
                sheet_combined = pd.concat([males, females], ignore_index=True)
        else:
            sheet_combined = pd.read_excel(master_path, sheet_name=sheet)
            if "Mouse Code" not in sheet_combined.columns:
                sheet_combined.rename(
                    columns={sheet_combined.columns[0]: "Mouse Code"}, inplace=True)

        sheet_combined = sheet_combined.dropna(subset=["Mouse Code"])
        sheet_combined["Progeny_Group"] = sheet
        all_compiled_data.append(sheet_combined)

    if not all_compiled_data:
        return pd.DataFrame()

    final_df = pd.concat(all_compiled_data, ignore_index=True)
    final_df[["Genotype", "Age_Extracted", "Sex_Extracted", "ID_Num"]] = \
        final_df["Mouse Code"].apply(lambda x: pd.Series(parse_mouse_code(x)))

    generate_grouped_tables(final_df, csv_out_dir, bone_target)
    return final_df


# ===========================================================
# PIPELINE EXECUTION
# ===========================================================

def execute_pipeline(data_folder, tibia_master, femur_master,
                     measurement_path, csv_out_dir, structure_type, fallback_bone):
    root_path = Path(os.path.normpath(data_folder))
    parse_anatomical_diameters(measurement_path, csv_out_dir)

    subfolders_with_data = {f.parent for f in root_path.rglob(FILE_GLOB_PATTERN)}
    if not subfolders_with_data:
        raise FileNotFoundError(
            f"No .txt files found in subfolders of {data_folder}")

    bone_groups = {"Tibia": [], "Femur": []}
    for folder in subfolders_with_data:
        assigned = determine_bone_type(folder, fallback_bone)
        bone_groups[assigned].append(folder)
        run_batch_bending_analysis(str(folder))

    all_analysis_files = list(root_path.rglob("Fz_Displacement_Analysis_*.xlsx"))
    if not all_analysis_files:
        raise FileNotFoundError(
            "Batch analysis produced no Excel summary files.")

    for bone_key, master_dest in [("Tibia", tibia_master), ("Femur", femur_master)]:
        if not bone_groups[bone_key]:
            continue

        sync_data_to_master(all_analysis_files, master_dest,
                            measurement_path, bone_key, fallback_bone, structure_type)

        try:
            compiled = process_all_sheets(master_dest, structure_type, csv_out_dir, bone_key)
            if not compiled.empty:
                csv_path = os.path.join(
                    os.path.dirname(master_dest),
                    f"FKBP5New_{bone_key}Master_Compiled.csv")
                compiled.to_csv(csv_path, index=False)
        except Exception as e:
            print(f"Post-processing error on {bone_key}: {e}")

    return True


# ===========================================================
# DASHBOARD ENTRY POINT
# ===========================================================

def run_workflow(inputs):
    print(f"Starting FKBP5 New Workflow for: {inputs.get('data_folder')}")
    return execute_pipeline(
        data_folder      = inputs.get("data_folder"),
        tibia_master     = inputs.get("tibia_master"),
        femur_master     = inputs.get("femur_master"),
        measurement_path = inputs.get("measurement_path"),
        csv_out_dir      = inputs.get("csv_out_dir"),
        structure_type   = inputs.get("structure_type"),
        fallback_bone    = inputs.get("fallback_bone"),
    )
