import os
import pandas as pd
import xlwings as xw

# Paths for the files
reference_file_path = r"Final Extraction.xlsx"
directory_path = r"Report Files"

# Validate directory
if not os.path.exists(directory_path):
    raise FileNotFoundError(f"Directory not found: {directory_path}")

# Step 1: Read the reference file (Column E) starting from E2 using xlwings
def read_reference_file(file_path):
    try:
        with xw.App(visible=False) as app:
            app.display_alerts = False
            app.screen_updating = False
            workbook = app.books.open(file_path, update_links=False, read_only=True)
            sheets = workbook.sheets

            # Debug: Print available sheet names
            print("Available sheets:", [sheet.name for sheet in sheets])

            sheet = sheets[0]  # Assuming the first sheet, adjust as needed
            print(f"Reading data from sheet: {sheet.name}")

            # Read all data in column E starting from E2 dynamically
            reference_values = sheet.range("E2").expand("down").value
            workbook.close()
    except Exception as e:
        raise ValueError(f"Error reading the Excel file: {e}")

    if reference_values is None:
        raise ValueError("No data found in the specified range (E2:E). Check the file.")

    return [str(value).strip() for value in reference_values if value is not None]

# Main script execution
try:
    expected_files = read_reference_file(reference_file_path)
except Exception as e:
    print(f"Failed to read the reference file: {e}")
    exit()

# Step 2: List all files in the directory (without extensions)
actual_files = os.listdir(directory_path)
actual_file_names = [os.path.splitext(file)[0] for file in actual_files]

# Debug: Print both lists to visually compare
print("Expected Files:")
print(expected_files)

print("\nActual Files:")
print(actual_file_names)

# Step 3: Preprocess filenames to normalize spaces and make case-insensitive comparisons
def clean_filename(filename):
    # Normalize spaces and make lowercase
    return " ".join(filename.split()).lower()

# Clean both expected and actual filenames
expected_files_clean = [clean_filename(f) for f in expected_files]
actual_files_clean = [clean_filename(f) for f in actual_file_names]

# Step 4: Identify missing files
missing_files = [
    file for file in expected_files if clean_filename(file) not in actual_files_clean
]

# Debug: Print missing files
print("\nMissing Files:")
for file in missing_files:
    print(repr(file))

# Step 5: Handle results
if not missing_files:
    print("No missing files detected.")
else:
    # Create a DataFrame for missing files
    missing_df = pd.DataFrame({"Missing File": missing_files})

    # Save the missing files report to an Excel file
    output_file_path = r"missing_files_report.xlsx"
    missing_df.to_excel(output_file_path, index=False)
    print(f"Missing files have been written to '{output_file_path}'.")
