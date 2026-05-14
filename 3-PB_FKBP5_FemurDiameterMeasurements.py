import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog


def select_file(title):
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
    root.destroy()
    return file_path


def select_directory(title):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    dir_path = filedialog.askdirectory(title=title)
    root.destroy()
    return dir_path


def process_bone_data(input_file, base_output_dir, bone_type="Femur"):
    if not input_file or not base_output_dir:
        print("Error: File or directory selection cancelled.")
        return

    xl = pd.ExcelFile(input_file)
    all_data = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)

        for index, row in df.iterrows():
            raw_code = str(row.iloc[0])

            if raw_code == 'nan' or bone_type not in raw_code:
                continue

            # Clean code and parse metadata
            clean_code = raw_code.split('_')[0]
            parts = clean_code.split('.')

            if len(parts) < 3:
                continue

            age_val = parts[1]
            sex_id = parts[2]
            sex_full = 'Male' if sex_id[0] == 'M' else 'Female'
            id_num = sex_id[1:]

            group_ce = row.iloc[2:5].astype(float)
            group_fh = row.iloc[5:8].astype(float)

            avg_ce = group_ce.mean()
            avg_fh = group_fh.mean()

            all_data.append({
                'ID_Num': id_num,
                'Age_Extracted': age_val,
                'Sex_Extracted': sex_full,
                'Progeny_Group': sheet,
                'Top Average': max(avg_ce, avg_fh),
                'Bottom Average': min(avg_ce, avg_fh)
            })

    master_df = pd.DataFrame(all_data)

    # These folders will be created inside the directory you click/select
    folder_top = os.path.join(base_output_dir, f"{bone_type}_Top_By_Genotype")
    folder_bottom = os.path.join(
        base_output_dir, f"{bone_type}_Bottom_By_Genotype")

    os.makedirs(folder_top, exist_ok=True)
    os.makedirs(folder_bottom, exist_ok=True)

    metrics = ['Top Average', 'Bottom Average']
    sort_priority = ['Wildtype', 'Mutant', 'Heterozygous']

    for metric in metrics:
        target_folder = folder_top if metric == 'Top Average' else folder_bottom
        for age in master_df['Age_Extracted'].unique():
            for sex in ['Male', 'Female']:
                subset = master_df[(master_df['Sex_Extracted'] == sex) &
                                   (master_df['Age_Extracted'] == age)].copy()
                if subset.empty:
                    continue

                table = subset.pivot_table(
                    index='ID_Num', columns='Progeny_Group', values=metric)

                def lineage_sort(col_name):
                    for i, gen in enumerate(sort_priority):
                        if col_name.startswith(gen):
                            return (i, col_name)
                    return (99, col_name)

                table = table[sorted(table.columns, key=lineage_sort)]
                file_name = f"{sex}_{age}Wks_{metric.replace(' ', '_')}.csv"
                table.to_csv(os.path.join(target_folder, file_name))

    print(f"\nSuccess! Processing complete for {bone_type}.")
    print(f"Results saved in: {base_output_dir}")


# --- INTERACTIVE SELECTION ---
print("Please select your Excel file...")
selected_input = select_file("Select Measurement Excel File")

print("Please select where you want the output folders to be created...")
selected_output = select_directory("Select Output Folder Location")

# Run for Femurs
process_bone_data(selected_input, selected_output, bone_type="Femur")
