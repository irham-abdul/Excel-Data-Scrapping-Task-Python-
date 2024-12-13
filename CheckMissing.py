import os
import pandas as pd

# Paths for the files
reference_file_path = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.xlsx"
directory_path = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Report Files"

# Step 1: Read the reference file (Column E) starting from E2 (skip header)
df_reference = pd.read_excel(reference_file_path, usecols=[4])  # Only read column E (index 4)
expected_files = (
    df_reference.iloc[1:, 0]  # Start from row 2 (E2)
    .dropna()  # Remove NaN values
    .astype(str)  # Ensure all values are strings
    .str.strip()  # Remove leading/trailing spaces
    .tolist()  # Convert to a Python list
)

# Step 2: List all files in the directory
actual_files = [os.path.splitext(file)[0] for file in os.listdir(directory_path)]  # Remove file extensions

# Step 3: Identify missing files by comparing "expected" and "actual"
missing_files = [file for file in expected_files if file not in actual_files]

# Step 4: Create a DataFrame for missing files and write it to a new Excel file
missing_df = pd.DataFrame(missing_files, columns=["Missing File"])

# Step 5: Save the result to a new Excel file
output_file_path = r"missing_files_report.xlsx"
missing_df.to_excel(output_file_path, index=False)

print(f"Missing files have been written to '{output_file_path}'.")
