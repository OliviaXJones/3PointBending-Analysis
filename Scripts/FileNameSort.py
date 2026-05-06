
INPUT_FOLDER = r"/Users/oliviajones/Desktop/OJ_ParsedBending"

file_times = []

for filename in os.listdir(INPUT_FOLDER):
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
