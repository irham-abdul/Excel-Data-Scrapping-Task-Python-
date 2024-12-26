import os
import pandas as pd

# Paths for the files
reference_file_path = r"C:\Users\mirham\OneDrive\WORK\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx"
directory_path = r"C:\Users\mirham\OneDrive\WORK\INTERN FILE\TASK 3\PART 2\Report Files"

# Validate directory
if not os.path.exists(directory_path):
    raise FileNotFoundError(f"Directory not found: {directory_path}")

# Step 1: Read the reference file (Column E) starting from E2 (skip header)
df_reference = pd.read_excel(reference_file_path, usecols=[4])  # Only read column E (index 4)
expected_files = (
    df_reference.iloc[0:, 0]  # Start from row 2 (E2)
    .dropna()  # Remove NaN values
    .astype(str)  # Ensure all values are strings
    .str.strip()  # Remove leading/trailing spaces
    .tolist()  # Convert to a Python list
)

# Step 2: List all files in the directory (without extensions)
actual_files = [os.path.splitext(file)[0] for file in os.listdir(directory_path)]

# Debug: Print both lists to visually compare
print("Expected Files:")
for file in expected_files:
    print(repr(file))

print("\nActual Files:")
for file in actual_files:
    print(repr(file))

# Step 3: Preprocess filenames to normalize spaces and make case-insensitive comparisons
def clean_filename(filename):
    # Normalize spaces and make lowercase
    return " ".join(filename.split()).lower()

# Clean both expected and actual filenames
expected_files_clean = [clean_filename(f) for f in expected_files]
actual_files_clean = [clean_filename(f) for f in actual_files]

# Step 4: Identify missing files
missing_files = [file for file in expected_files if clean_filename(file) not in actual_files_clean]

# Debug: Print missing files
print("\nMissing Files:")
for file in missing_files:
    print(repr(file))

# Step 5: Handle results
if not missing_files:
    print("No missing files detected.")
else:
    missing_df = pd.DataFrame(missing_files, columns=["Missing File"])
    output_file_path = r"missing_files_report.xlsx"
    missing_df.to_excel(output_file_path, index=False)
    print(f"Missing files have been written to '{output_file_path}'.")
