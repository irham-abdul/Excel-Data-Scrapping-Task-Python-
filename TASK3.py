import os
from openpyxl.styles import Font
from datetime import datetime, timedelta
import pandas as pd                             # pip install pandas
from openpyxl import Workbook, load_workbook    # pip install openpyxl
import win32com.client as win32                 # pip install pywin32

def convert_xlsb_to_xlsx(xlsb_file, xlsx_file):
    #Convert a .xlsb file to a .xlsx file using win32com.client
    try:
        # Create a COM object for Excel
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # FALSE to hide Excel application window
        
        wb = excel.Workbooks.Open(xlsb_file) # Open the .xlsb file
        
        # Save the .xlsb as .xlsx
        wb.SaveAs(xlsx_file, FileFormat=51)  # FileFormat=51 is to .xlsx
        wb.Close() 
        excel.Quit()  
        return True  # Return True when conversion is successful
    except Exception as e:
        print(f"Error converting {xlsb_file} to {xlsx_file}: {e}")
        if 'excel' in locals():
            excel.Quit()  # Quit Excel if it was opened
        return False  # Return False if error occurs

def extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_column):
    try:
        # Check the file extension to decide method for reading the file
        file_extension = os.path.splitext(source_file)[1].lower()

        # If file is .xlsb, convert it to .xlsx before extracting data
        if file_extension == '.xlsb':
            xlsx_file = source_file.replace('.xlsb', '.xlsx')
            if not convert_xlsb_to_xlsx(source_file, xlsx_file):
                return False  # Return False if conversion fails
            source_file = xlsx_file  # Update the source file to the converted .xlsx file

        # Reading the .xlsx file (both new and converted)
        df_source = pd.read_excel(source_file, sheet_name=source_sheet_name, index_col=None)
        source_wb = load_workbook(source_file)

        # Get the value from cell B2 in the first sheet of .xlsx
        value_b2 = source_wb.worksheets[0]['B2'].value  # Accessing the first sheet (index 0)

        if value_b2 is None:
            value_b2 = "No value in B2"
        
        # Get name of the report from the source file without the extension
        report_name = os.path.basename(source_file).split('.')[0]  

        # Get the absolute path of the source file
        source_file_path = os.path.abspath(source_file)
        
        # Load the target workbook and worksheet
        target_wb = load_workbook(target_file)
        target_ws = target_wb[target_sheet_name]
        
        # Identify the last used row in the target sheet
        last_row = target_ws.max_row  # max_row returns the last used row in the sheet
        
        # Make the row where new data should start (after the last row with data)
        target_row_start = last_row + 1
        
        bold_font = Font(bold=True)
        target_ws.cell(row=1, column=12).font = bold_font

        # Insert the report name into column B 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=2, value=report_name)  
        
        # Insert the source file path into column C 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=3, value=source_file_path)  
        
        # Insert the value from B2 of the first sheet into column D 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=4, value=value_b2)  

        # Calculate yesterday's date in the required format (dd/mm/yyyy)
        yesterday_date = (datetime.now() - timedelta(1)).strftime('%d/%m/%Y')

        # Insert "zsystem" header and yesterday's date in the first row only (column L)
        if target_row_start == last_row + 1:  # Only for the first row written
            target_ws.cell(row=1, column=12, value="zsystem").font = bold_font  # Add header in L1 (bold font)
            target_ws.cell(row=2, column=12, value=yesterday_date)  # Add date in L2
        
        # Loop through each row in the source, starting from the specified start index
        for row_offset in range(len(df_source)):
            row_data = df_source.iloc[row_offset, 1:8] 
            
            target_row = target_row_start + row_offset
            # Loop through columns to populate data
            for col_offset, value in enumerate(row_data):
                target_ws.cell(row=target_row, column=5 + col_offset, value=value)  

            # Insert CONCATENATE formula into Column F for current row, dynamically using value from column 5 (E)
            target_ws.cell(
                row=target_row, 
                column=6,  # Column F is column 6 (1-based index)
                value=f'=CONCATENATE("{target_ws.cell(row=target_row, column=5).value} as at ","-",TEXT(L2,"dd-mmm-yyyy"))'
            )

        # Save updated target workbook
        target_wb.save(target_file)
        return True  # Return True if data extraction and appending are successful

    except Exception as e:
        print(f"An error occurred with {source_file}: {str(e)}")
        return False  # Return False when an error occurs during extraction

# Function to read file paths from the text file
def get_source_files(file_path):
    try:
        with open(file_path, 'r') as f:
            # Read all lines from the text file and strip any extra spaces or newlines
            file_paths = [line.strip() for line in f.readlines()]
        return file_paths
    except FileNotFoundError:
        print(f"Text file not found: {file_path}")
        return []

# File paths, sheet names, and row/column details for target
source_files = get_source_files(r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Report_Template_Paths.txt") # Read the source files from the text file
target_file = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx"
source_sheet_name = "Rpt-Maintain"  
target_sheet_name = "Sheet1"         
source_row_start_index = 0           # Start at the first row in the source sheet, Base 0
target_column = 4                    # Starting column to insert into in the target sheet, 1-based index

# Initialize a list to track successful extractions
successful_extractions = []

# Process each source file in the list
for idx, source_file in enumerate(source_files):
    if extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_column):
        successful_extractions.append(f"Finished extracting data from report {idx + 1}")
    else:
        successful_extractions.append(f"Failed to extract data from report {idx + 1}")

# Print summary after all reports have been processed
print("\n".join(successful_extractions))

# After processing, output the final extraction data to text and csv
df = pd.read_excel(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx', sheet_name='Sheet1')  # target file to write extracted data
df.to_csv(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.txt', sep='\t', index=False)  # Portion to write from excel file to text.
print("Successfully output excel data to Final_Extraction.txt") 
df.to_csv(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.csv', sep='\t', index=False)  # Portion to write from excel file to csv.
print("Successfully output excel data to Final_Extraction.csv") 
