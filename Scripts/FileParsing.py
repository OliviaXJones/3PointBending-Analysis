import os
import re
from datetime import datetime

INPUT_FOLDER = r"/Users/oliviajones/Desktop/Torsion"
OUTPUT_FOLDER = r"/Users/oliviajones/Desktop/Torsion"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

for filename in os.listdir(INPUT_FOLDER):
    if not filename.endswith(".txt"):
        continue
    filepath = os.path.join(INPUT_FOLDER, filename)
    
    with open(filepath, 'r') as f:
        lines = f.readlines()

    current_block = []
    timestamp = None
    move_type = None
    test_count = 0
    in_test = False

    for line in lines:
        stripped = line.strip()

        # Start of new test
        if stripped == "<INFO>":
            current_block = [line]
            timestamp = None
            move_type = None
            in_test = True
            continue

        if in_test:
            current_block.append(line)

            # Capture Date/Time for naming
            if stripped.startswith("Date:"):
                date_str = stripped.split("Date:")[1].strip().replace(",", "").replace(" ", "_")
            if stripped.startswith("Time:"):
                time_str = stripped.split("Time:")[1].strip().replace(":", "-")
                if date_str:
                    timestamp = f"{date_str}_{time_str}"

            # Detect move type
            if "<Move Relative>" in stripped:
                move_type = "Relative"

            # End of data block
            if stripped == "<END DATA>":
                in_test = False
                # Only save if it's Move Relative
                if move_type == "Relative":
                    test_count += 1
                    name_parts = [os.path.splitext(filename)[0], f"test{test_count}", move_type]
                    if timestamp:
                        name_parts.append(timestamp)
                    out_name = "_".join(name_parts) + ".txt"
                    out_path = os.path.join(OUTPUT_FOLDER, out_name)

                    with open(out_path, 'w') as out_f:
                        out_f.writelines(current_block)
                    print(f"✅ Saved {out_name}")

##################################
# MECHANICAL TESTING SORTED LIST #
##################################

file_times = []

for filename in os.listdir(OUTPUT_FOLDER):
    if not filename.endswith(".txt"):
        continue
    
    # Extract only the date/time part at the end of the filename
    # Example: Humerus12wk6_test2_Relative_Mon_Nov_10_2025_13-16-53.915
    match = re.search(r'_(Mon|Tue|Wed|Thu|Fri|Sat|Sun)_(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)_\d{1,2}_\d{4}_\d{2}-\d{2}-\d{2}\.\d+', filename)
    if match:
        datetime_str = match.group(0).lstrip('_')  # remove leading underscore
        # Parse datetime
        try:
            dt = datetime.strptime(datetime_str, "%a_%b_%d_%Y_%H-%M-%S.%f")
            file_times.append((dt, filename))
        except Exception as e:
            print(f"⚠️ Could not parse datetime for {filename}: {e}")
    else:
        print(f"⚠️ No timestamp found in {filename}")

# Sort files by datetime
file_times.sort(key=lambda x: x[0])

# Print sorted list
print("📋 Sorted file list by date/time:")
for i, (dt, fname) in enumerate(file_times, start=1):
    print(f"{i:02d}: {fname} ({dt})")

# Optional: save to text file
with open(os.path.join(INPUT_FOLDER, "sorted_file_list.txt"), "w") as f:
    for dt, fname in file_times:
        f.write(f"{fname}\t{dt}\n")

print("✅ Sorted list saved as sorted_file_list.txt")

