import pandas as pd

# --- USER INPUT ---
file1 = "/Users/oliviajones/Desktop/SHP099_Medigel_111925/Fz_Displacement_Analysis.xlsx"    # path to first Excel file
file2 = "/Users/oliviajones/Documents/SHP_Medigel.xlsx"    # path to second Excel file

sheet1 = "Sheet1"           # sheet in first file
sheet2 = "Sheet1"           # sheet in second file

col1 = "Energy_to_Failure_Nmm"            # column in first sheet
col2 = "Work-to-fracture (Nmm)"            # column in second sheet
# -------------------

# Load the sheets
df1 = pd.read_excel(file1, sheet_name=sheet1)
df2 = pd.read_excel(file2, sheet_name=sheet2)

# Ensure same comparison length
min_len = min(len(df1), len(df2))

# Compare values
comparison = df1[col1].iloc[:min_len] == df2[col2].iloc[:min_len]

# Summary
total = len(comparison)
matches = comparison.sum()
mismatches = total - matches

print(f"Total compared rows: {total}")
print(f"Matches: {matches}")
print(f"Mismatches: {mismatches}")

# Show mismatched rows
print("\nMismatched rows:")
mismatch_df = pd.DataFrame({
    f"{file1}_{col1}": df1[col1].iloc[:min_len].values,
    f"{file2}_{col2}": df2[col2].iloc[:min_len].values,
    "Match": comparison.values
})

print(mismatch_df[mismatch_df["Match"] == False])
