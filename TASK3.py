import pandas as pd  # Import library first (pip install)
from openpyxl import load_workbook
from pyxlsb import open_workbook as open_workbook_b
import os
from datetime import datetime, timedelta
from openpyxl.styles import Font

def extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_column):
    try:
        # Check the file extension to decide the method for reading the file
        file_extension = os.path.splitext(source_file)[1].lower()
        
        # Load source Excel file based on its extension
        if file_extension == '.xlsx':
            df_source = pd.read_excel(source_file, sheet_name=source_sheet_name, index_col=None)
            source_wb = load_workbook(source_file)
        elif file_extension == '.xlsb':
            with open_workbook_b(source_file) as wb:
                sheet = wb.get_sheet(source_sheet_name)
                # Convert the sheet to a DataFrame
                df_source = pd.DataFrame(sheet.rows())
            source_wb = None  # Since we don't need openpyxl for .xlsb files
        else:
            raise ValueError("Unsupported file type. Only .xlsx and .xlsb are supported.")
        
        # Get name of the report from the source file without the extension
        report_name = os.path.basename(source_file).split('.')[0]  
        
        # Get the absolute path of the source file
        source_file_path = os.path.abspath(source_file)
        
        # If source_wb is available (for .xlsx files), use openpyxl to get the value from B2
        if source_wb:
            source_ws = source_wb[source_wb.sheetnames[0]]  # Access the first sheet, base index 0
            value_b2 = source_ws['B2'].value
        else:
            value_b2 = None  # For .xlsb, B2 value extraction can be omitted or handled differently

        print(f"Value from B2: {value_b2}")  
        
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
        # All starting from the identified row (target_row_start)
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=2, value=report_name)  
        
        # Insert the source file path into column C 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=3, value=source_file_path)  
        
        # Insert the value from B2 of the first sheet into column D 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=4, value=value_b2)  

        # Calculate yesterday's date
        yesterday_date = (datetime.now() - timedelta(1)).strftime('%Y-%m-%d')
        
        # Set the header for the 'zsystem' column (column 11)
        target_ws.cell(row=target_row_start - 2, column=12, value='zsystem')

        # Loop through each row in the source (Sheet2), starting from the specified start index
        for row_offset in range(len(df_source)):
            row_data = df_source.iloc[row_offset, 1:8] 
            
            target_row = target_row_start + row_offset
            
            for col_offset, value in enumerate(row_data):
                target_ws.cell(row=target_row, column=5 + col_offset, value=value)  

            # Insert CONCATENATE formula into Column F for current row, dynamically using value from column 5 (E)
            target_ws.cell(
                row=target_row, 
                column=6,  # Column F is column 6 (1-based index)
                value=f'=CONCATENATE("{target_ws.cell(row=target_row, column=5).value} as at ","-",TEXT(L2,"dd-mmm-yyyy"))'
            )

            # Insert yesterday's date into the 'zsystem' column (column 11) for the first data row (row_offset == 0)
            if row_offset == 0:  # Only insert in the first row of data
                target_ws.cell(row=target_row - 1, column=12, value=yesterday_date)  # Column 11 for 'zsystem'
                
        # Save updated target workbook
        target_wb.save(target_file)
        print(f"Rows appended for {source_file}, report name added to column B, source file path added to column C, value from B2 added to column D, and data from Columns 1-7 added to Columns E-K with formula in Column F, and 'zsystem' added to Column 11.")
    
    except FileNotFoundError:
        print(f"File not found: {source_file} or {target_file}")
    except KeyError:
        print(f"Sheet '{source_sheet_name}' or '{target_sheet_name}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


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

# Process each source file in the list
for source_file in source_files:
    extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_column)

df = pd.read_excel(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx', sheet_name='Sheet1')   # target file to write extracted data
df.to_csv(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.txt', sep='\t', index=False)           # Portion to write from excel file to text.
print("Successfully output excel data to Final_Extraction.txt") 
df.to_csv(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.csv', sep='\t', index=False)           # Portion to write from excel file to csv.
print("Successfully output excel data to Final_Extraction.csv") 
