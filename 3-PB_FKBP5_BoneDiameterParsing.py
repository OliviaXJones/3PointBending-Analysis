import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox


def show_step_instruction(title, message):
    """Shows a popup with clear instructions before an action."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo(title, message)
    root.destroy()


def select_file(instruction_title, explorer_title):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    # The instruction title helps the user know what's coming
    file_path = filedialog.askopenfilename(
        title=explorer_title,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    root.destroy()
    return file_path


def select_directory(explorer_title):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    dir_path = filedialog.askdirectory(title=explorer_title)
    root.destroy()
    return dir_path


def process_all_bones(input_file, base_output_dir):
    if not input_file or not base_output_dir:
        show_step_instruction(
            "Process Cancelled", "You didn't select a file or folder, so the process stopped.")
        return

    try:
        xl = pd.ExcelFile(input_file)

        for bone_type in ["Femur", "Tibia"]:
            if bone_type == "Femur":
                label_top, label_bottom = "Anteroposterior_Femur", "Mediolateral_Femur"
            else:
                label_top, label_bottom = "Proximal_Tibia", "Distal_Tibia"

            all_data = []

            for sheet in xl.sheet_names:
                df = pd.read_excel(input_file, sheet_name=sheet)

                for index, row in df.iterrows():
                    raw_code = str(row.iloc[0])
                    if raw_code == 'nan':
                        continue

                    if bone_type == "Femur" and "_Femur" not in raw_code:
                        continue
                    if bone_type == "Tibia" and "_Femur" in raw_code:
                        continue

                    clean_code = raw_code.split('_')[0]
                    parts = clean_code.split('.')
                    if len(parts) < 3:
                        continue

                    age_val, sex_id = parts[1], parts[2]
                    sex_full = 'Male' if sex_id[0] == 'M' else 'Female'
                    id_num = sex_id[1:]

                    group_ce = row.iloc[2:5].astype(float)
                    group_fh = row.iloc[5:8].astype(float)
                    avg_ce, avg_fh = group_ce.mean(), group_fh.mean()

                    all_data.append({
                        'ID_Num': id_num,
                        'Age_Extracted': age_val,
                        'Sex_Extracted': sex_full,
                        'Progeny_Group': sheet,
                        'Top_Val': max(avg_ce, avg_fh),
                        'Bottom_Val': min(avg_ce, avg_fh)
                    })

            if not all_data:
                continue

            master_df = pd.DataFrame(all_data)
            folder_top = os.path.join(base_output_dir, label_top)
            folder_bottom = os.path.join(base_output_dir, label_bottom)

            os.makedirs(folder_top, exist_ok=True)
            os.makedirs(folder_bottom, exist_ok=True)

            mapping = [('Top_Val', folder_top, label_top),
                       ('Bottom_Val', folder_bottom, label_bottom)]

            sort_priority = ['Wildtype', 'Mutant', 'Heterozygous']

            for data_key, target_folder, file_label in mapping:
                for age in master_df['Age_Extracted'].unique():
                    for sex in ['Male', 'Female']:
                        subset = master_df[(master_df['Sex_Extracted'] == sex) &
                                           (master_df['Age_Extracted'] == age)].copy()
                        if subset.empty:
                            continue

                        table = subset.pivot_table(
                            index='ID_Num', columns='Progeny_Group', values=data_key)

                        def lineage_sort(col_name):
                            for i, gen in enumerate(sort_priority):
                                if col_name.startswith(gen):
                                    return (i, col_name)
                            return (99, col_name)

                        table = table[sorted(table.columns, key=lineage_sort)]
                        file_name = f"{sex}_{age}Wks_{file_label}.csv"
                        table.to_csv(os.path.join(target_folder, file_name))

        show_step_instruction(
            "Success!", f"All done! Your folders have been created in:\n{base_output_dir}")

    except Exception as e:
        show_step_instruction(
            "Error", f"Something went wrong while reading the file:\n{str(e)}")

# --- THE ACTUAL USER EXPERIENCE FLOW ---


# STEP 1: Explaining the Excel selection
show_step_instruction("Step 1: The Input File",
                      "Next, a window will open. Please navigate to and select the EXCEL FILE that contains your measurements.")

file_to_open = select_file(
    "Select Excel File", "Choose your Measurements Excel file (.xlsx)")

# STEP 2: Explaining the Folder selection
if file_to_open:
    show_step_instruction("Step 2: The Saving Location",
                          "Great! Now, choose the FOLDER where you want the new anatomical CSV folders to be created.")

    folder_to_save = select_directory("Choose where to save the results")

    # STEP 3: Running the magic
    if folder_to_save:
        process_all_bones(file_to_open, folder_to_save)
