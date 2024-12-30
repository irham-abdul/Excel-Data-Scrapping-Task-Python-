import os
from datetime import datetime
import pandas as pd

def read_directories_from_file(file_path):
    """Read directory paths from a text file."""
    try:
        with open(file_path, 'r') as file:
            directories = [line.strip() for line in file if line.strip()]
        return directories
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        return []

def get_file_metadata(directory):
    """Traverse directory and collect file metadata."""
    file_data = []

    for root, _, files in os.walk(directory):
        for file in files:
            filepath = os.path.join(root, file)  # Full file path
            try: 
                file_size = os.path.getsize(filepath)
                modified_time = os.path.getmtime(filepath)
                formatted_time = datetime.fromtimestamp(modified_time).strftime('%Y-%m-%d %H:%M:%S')

                # Append metadata with only directory path for "File Path"
                file_data.append({
                    "Directory": directory,
                    "File Name": file,
                    "File Path": root,  # Directory path only
                    "File Size (Bytes)": file_size,
                    "Last Modified": formatted_time
                })
            except Exception as e:
                print(f"Error accessing file {filepath}: {e}")
    return file_data

def save_to_excel(file_data, output_file):
    """Save collected metadata into an Excel file."""
    df = pd.DataFrame(file_data)
    df.to_excel(output_file, index=False)
    print(f"Metadata saved to {output_file}")

if __name__ == "__main__":
    directories_file = "Report_Template_Paths.txt"  # Text file with directories
    output_file = "File_List_w_Metadata.xlsx"      # Output Excel file name

    directories = read_directories_from_file(directories_file)

    if directories:
        all_file_metadata = []
        for directory in directories:
            print(f"Scanning directory: {directory}")
            all_file_metadata.extend(get_file_metadata(directory))
        if all_file_metadata:
            save_to_excel(all_file_metadata, output_file)
        else:
            print("No files found in the specified directories.")
    else:
        print("No directories to scan.")
