import pandas as pd
import os

input_path = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending\Measurement Files\FKBP5Null_Tibia+Femur_11226.xlsx"
output_directory = r"C:\Users\olivi\OneDrive - Medical University of South Carolina\3-Point Bending"


def process_tibia_with_grouping(input_file, base_dir):
    xl = pd.ExcelFile(input_file)
    all_data = []

    # 1. Processing all sheets and extracting metadata from the code
    for sheet in xl.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)

        for index, row in df.iterrows():
            mouse_code = str(row.iloc[0])
            if '_Femur' in mouse_code or mouse_code == 'nan':
                continue

            # Parsing "Genotype.Age.Sex&ID" (e.g., M.12.M24)
            # We split by the dots
            parts = mouse_code.split('.')
            if len(parts) < 3:
                continue

            genotype_letter = parts[0]
            age_val = parts[1]
            sex_id = parts[2]  # e.g., M24

            # Extract Sex and ID from the last part
            sex_letter = sex_id[0]
            sex_full = 'Male' if sex_letter == 'M' else 'Female'
            id_num = sex_id[1:]

            # Calculate Averages (Cols C-E and F-H)
            group_ce = row.iloc[2:5].astype(float)
            group_fh = row.iloc[5:8].astype(float)

            avg_ce = group_ce.mean()
            avg_fh = group_fh.mean()

            top_avg = max(avg_ce, avg_fh)
            bottom_avg = min(avg_ce, avg_fh)

            all_data.append({
                'ID_Num': id_num,
                'Age_Extracted': age_val,
                'Sex_Extracted': sex_full,
                # Original sheet name (Wildtype, Mutant F0, etc.)
                'Progeny_Group': sheet,
                'Top Average': top_avg,
                'Bottom Average': bottom_avg
            })

    # Convert everything to one big master DataFrame
    master_df = pd.DataFrame(all_data)

    # 2. Setup Folders for CSV output
    folder_top = os.path.join(base_dir, "Tibia_Top_By_Genotype")
    folder_bottom = os.path.join(base_dir, "Tibia_Bottom_By_Genotype")
    os.makedirs(folder_top, exist_ok=True)
    os.makedirs(folder_bottom, exist_ok=True)

    # 3. Generate the grouped .csv files
    metrics = ['Top Average', 'Bottom Average']
    sort_priority = ['Wildtype', 'Mutant', 'Heterozygous']
    unique_ages = master_df['Age_Extracted'].unique()

    for metric in metrics:
        target_folder = folder_top if metric == 'Top Average' else folder_bottom

        for age in unique_ages:
            for sex in ['Male', 'Female']:
                # Filter data for this specific age and sex
                subset = master_df[(master_df['Sex_Extracted'] == sex) &
                                   (master_df['Age_Extracted'] == age)].copy()

                if subset.empty:
                    continue

                # Create pivot table: Rows = ID, Columns = Genotype/Sheet
                table = subset.pivot_table(index='ID_Num',
                                           columns='Progeny_Group',
                                           values=metric)

                # Sort columns based on your WT -> Mutant -> Het priority
                def lineage_sort(col_name):
                    for i, gen in enumerate(sort_priority):
                        if col_name.startswith(gen):
                            return (i, col_name)
                    return (99, col_name)

                sorted_cols = sorted(table.columns, key=lineage_sort)
                table = table[sorted_cols]

                # Save the CSV
                file_name = f"{sex}_{age}Wks_{metric.replace(' ', '_')}.csv"
                table.to_csv(os.path.join(target_folder, file_name))

    print("Success! CSV files organized by Age and Sex are ready.")


process_tibia_with_grouping(input_path, output_directory)
