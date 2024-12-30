import os
import pandas as pd

# Define the base directory where folders will be created
base_dir = '/Users/muhammadirham/Library/CloudStorage/OneDrive-Personal/WORK/INTERN FILE/TASK 3/PART 2/SAMPLE FILES'

# Ensure the base directory exists, if not, create it
if not os.path.exists(base_dir):
    os.makedirs(base_dir)

# Number of folders and files
num_folders = 5
num_files = 5

# Create folders and files
for i in range(1, num_folders + 1):
    folder_name = f"Folder_{i}"
    folder_path = os.path.join(base_dir, folder_name)
    
    # Create a folder if it doesn't already exist
    os.makedirs(folder_path, exist_ok=True)
    
    for j in range(1, num_files + 1):
        file_name = f"File_{j}.xlsx"
        file_path = os.path.join(folder_path, file_name)
        
        # Create a dummy DataFrame
        df = pd.DataFrame({
            'Column1': range(100),
            'Column2': range(100, 200)
        })
        
        # Write the DataFrame to an Excel file
        df.to_excel(file_path, index=False)

print("Folders and files have been created successfully.")
