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
actual_files = os.listdir(directory_path)
actual_files_metadata = []

for file in actual_files:
    file_path = os.path.join(directory_path, file)
    file_name, _ = os.path.splitext(file)
    metadata = {
        "file_name": file_name,
        "modified_date": os.path.getmtime(file_path),  # Last modified timestamp
        "file_size": os.path.getsize(file_path),  # File size in bytes
    }
    actual_files_metadata.append(metadata)

# Convert modified timestamp to readable format
for metadata in actual_files_metadata:
    metadata["modified_date"] = pd.to_datetime(metadata["modified_date"], unit="s")

# Convert metadata to DataFrame for easy processing
df_actual_files = pd.DataFrame(actual_files_metadata)

# Debug: Print both lists to visually compare
print("Expected Files:")
for file in expected_files:
    print(repr(file))

print("\nActual Files with Metadata:")
print(df_actual_files)

# Step 3: Preprocess filenames to normalize spaces and make case-insensitive comparisons
def clean_filename(filename):
    # Normalize spaces and make lowercase
    return " ".join(filename.split()).lower()

# Clean both expected and actual filenames
expected_files_clean = [clean_filename(f) for f in expected_files]
df_actual_files["cleaned_file_name"] = df_actual_files["file_name"].apply(clean_filename)

# Step 4: Identify missing files
missing_files = [
    file for file in expected_files if clean_filename(file) not in df_actual_files["cleaned_file_name"].tolist()
]

# Debug: Print missing files
print("\nMissing Files:")
for file in missing_files:
    print(repr(file))

# Step 5: Handle results
if not missing_files:
    print("No missing files detected.")
else:
    # Add metadata for missing files (if available)
    missing_files_metadata = []
    for file in missing_files:
        clean_file = clean_filename(file)
        metadata = df_actual_files[df_actual_files["cleaned_file_name"] == clean_file]
        if not metadata.empty:
            # Include only relevant columns
            missing_files_metadata.append({
                "Missing File": file,
                "Modified Date": metadata.iloc[0]["modified_date"],
                "File Size (Bytes)": metadata.iloc[0]["file_size"],
            })
        else:
            # If metadata is not available, add placeholders
            missing_files_metadata.append({
                "Missing File": file,
                "Modified Date": None,
                "File Size (Bytes)": None,
            })
    
    # Create a DataFrame for missing files with metadata
    missing_df = pd.DataFrame(missing_files_metadata)
    
    # Save the missing files report to an Excel file
    output_file_path = r"missing_files_report_with_metadata.xlsx"
    missing_df.to_excel(output_file_path, index=False)
    print(f"Missing files have been written to '{output_file_path}'.")
