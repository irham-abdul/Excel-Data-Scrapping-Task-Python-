import os
import openpyxl
import xlwings as xw
from datetime import datetime, timedelta
import psutil  # Module for managing system processes

# Path to the text file containing the list of Excel files and folders
txt_file_path = r"C:\Users\mirham\OneDrive\WORK\INTERN FILE\TASK 3\PART 2\Report_Template_Paths.txt"
# Path to the destination file
destination_file_path = r"C:\Users\mirham\OneDrive\WORK\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx"

# Function to terminate any open Excel processes
def close_open_excel_files():
    print("Closing any open Excel files...")
    for process in psutil.process_iter(attrs=['name']):
        if process.info['name'] and process.info['name'].lower() == 'excel.exe':
            try:
                process.terminate()  # Terminate the process
                process.wait(timeout=5)  # Wait for it to close
                print(f"Terminated: {process.info['name']}")
            except Exception as e:
                print(f"Failed to terminate {process.info['name']}: {e}")

# Function to convert .xlsb to .xlsx
def convert_xlsb_to_xlsx(xlsb_file_path):
    # Specify the directory where the converted files should be saved
    converted_dir = r"C:\Users\mirham\OneDrive\WORK\INTERN FILE\TASK 3\PART 2\Converted"
    
    # Ensure the directory exists
    if not os.path.exists(converted_dir):
        os.makedirs(converted_dir)
    
    # Generate the new file path for the converted xlsx file
    base_name = os.path.basename(xlsb_file_path)  # Extract file name from the full path
    xlsx_file_name = base_name.replace('.xlsb', '.xlsx')  # Replace extension
    xlsx_file_path = os.path.join(converted_dir, xlsx_file_name)  # Construct full path
    
    # Open the .xlsb file using xlwings in the background (without GUI)
    with xw.App(visible=False) as app:
        # Open the .xlsb file
        wb = app.books.open(xlsb_file_path)
        app.enable_events = False
        # Save it as .xlsx in the converted directory
        wb.save(xlsx_file_path)
        
        # Close the workbook
        wb.close()

    return xlsx_file_path

def process_files(file_paths):
    # Check if the destination file exists
    if not os.path.exists(destination_file_path):
        # If it doesn't exist, create a new workbook and save it first
        print(f"Creating new destination file at {destination_file_path}")
        destination_wb = openpyxl.Workbook()  # Create a new workbook
        destination_ws = destination_wb.active
    else:
        # If the file exists, open it
        print(f"Opening existing destination file at {destination_file_path}")
        destination_wb = openpyxl.load_workbook(destination_file_path)
        destination_ws = destination_wb.active

    # Clear all data (including the old header) in the sheet
    destination_ws.delete_rows(1, destination_ws.max_row)

    # Create the new header in row 1 (removed the "No" column)
    header = ["Report Template", "Template Path", "Report Path", "Report Name", "Origin Value", 
              "Filter", "Text", "Report Format", "Frequency", "Term", "zsystem"]
    
    # Add the header to row 1 and make it bold
    destination_ws.append(header)
    for cell in destination_ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)

    # Process each file listed in the text file
    for file_path in file_paths:
        file_path = file_path.strip()  # Remove any extra whitespace or newline characters
        if os.path.isdir(file_path):  # If it's a folder, process all Excel files inside it
            for root, dirs, files in os.walk(file_path):
                for file in files:
                    if file.endswith('.xlsx') or file.endswith('.xlsb'):
                        full_file_path = os.path.join(root, file)
                        print(f"Processing file: {full_file_path}")
                        process_excel_file(full_file_path, destination_ws)
        elif os.path.exists(file_path):  # If it's an individual file, process it directly
            print(f"Processing file: {file_path}")
            process_excel_file(file_path, destination_ws)
        else:
            print(f"File or folder does not exist: {file_path}")

    # Save the changes to the destination file
    destination_wb.save(destination_file_path)
    print("Data extraction complete.")

# Function to process an individual Excel file
def process_excel_file(file_path, destination_ws):
    if file_path.endswith('.xlsb'):
        file_path = convert_xlsb_to_xlsx(file_path)

    # Open the source Excel file with data_only=True to get evaluated formulas
    content_wb = openpyxl.load_workbook(file_path, data_only=True)
    
    # Get Sheet 1 (index 0) and Sheet 2 (index 1)
    sheet_1 = content_wb.worksheets[0]
    sheet_2 = content_wb.worksheets[1]

    # Get the value from B2 in Sheet 1 (for Column D in destination file)
    value_from_b2 = sheet_1['B2'].value

    # Get the number of rows in Sheet 1 and Sheet 2 (we'll use this for filling Column D, E, F, etc.)
    rows_count_sheet_1 = sheet_1.max_row
    rows_count_sheet_2 = sheet_2.max_row

    # Calculate the previous day's date
    previous_date = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")

    # Iterate through the rows of the source file and append data to the destination file
    for i in range(2, max(rows_count_sheet_1, rows_count_sheet_2) + 1):
        # Column A: Sequential numbers are removed now
        row_data = []

        # Column B: Extract the file name (excluding the extension)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        row_data.append(file_name)

        # Column C: Extract the file path
        row_data.append(file_path)

        # Column D: Fill with the value from B2 of Sheet 1 (same value for every row in Column D)
        row_data.append(value_from_b2)

        # Columns G to J: Extract data from Columns D to H of Sheet 2
        for col_letter in ['C', 'D', 'E', 'F', 'G', 'H']:
            if i <= rows_count_sheet_2:
                row_data.append(sheet_2[f"{col_letter}{i}"].value)
            else:
                row_data.append(None)

        # Append the row data to the destination file
        destination_ws.append(row_data)

    # Add the previous day's date to cell K2
    destination_ws['K2'] = previous_date

# Close any open Excel files before starting the process
close_open_excel_files()

# Read file paths from the text file
with open(txt_file_path, "r") as file:
    file_paths = file.readlines()

# Process the files (both individual files and files inside folders)
process_files(file_paths)
