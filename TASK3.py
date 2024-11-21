import pandas as pd #import library first(pip install)
from openpyxl import load_workbook
import os
#nkhiyhxguextgvjv
print("try")
def extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_row_start, target_column):
    try:
        # Load the source Excel file as a DataFrame (reading Sheet2 for the data)
        df_source = pd.read_excel(source_file, sheet_name=source_sheet_name, index_col=None)
        
        # Get the name of the report from the source file (without the extension)
        report_name = os.path.basename(source_file).split('.')[0]  # Get the file name without extension
        
        # Get the absolute path of the source file
        source_file_path = os.path.abspath(source_file)  # Full path to the source file
        
        # Load the source workbook and reference the first sheet
        source_wb = load_workbook(source_file)
        source_ws = source_wb[source_wb.sheetnames[0]]  # Access the first sheet
        
        # Extract the value from cell B2 of the first sheet
        value_b2 = source_ws['B2'].value  # Get the value from cell B2
        print(f"Value from B2: {value_b2}")  # Debug print statement to check the value
        
        # Load the target workbook and worksheet
        target_wb = load_workbook(target_file)
        target_ws = target_wb[target_sheet_name]
        
        # Insert the report name into column B starting at row 3
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=2, value=report_name)  # Column B is column 2 (1-based index)
        
        # Insert the source file path into column C starting at row 3
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=3, value=source_file_path)  # Column C is column 3 (1-based index)
        
        # Insert the value from B2 of the first sheet into column D starting at row 3
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=4, value=value_b2)  # Column D is column 4 (1-based index)

        # Loop through each row in the source (Sheet2), starting from the specified start index
        for row_offset in range(len(df_source)):
            # Access the row data from columns 1 to 7 (2nd to 8th columns in Excel)
            row_data = df_source.iloc[row_offset, 1:8]  # Columns 1 to 7 are columns 2 to 8 in Excel (0-based indexing)
            
            # Set the target row position in the target sheet, adjusted to 1-based indexing
            target_row = target_row_start + row_offset
            
            # Insert each value from row_data (columns 1 to 7) into the target sheet starting at column 5 (Column E)
            for col_offset, value in enumerate(row_data):
                target_ws.cell(row=target_row, column=5 + col_offset, value=value)  # Insert into Column E to K (5 to 11)

            # Insert the CONCATENATE formula into Column F for the current row
            target_ws.cell(
                row=target_row, 
                column=6,  # Column F is column 6 (1-based index)
                value=f'=CONCATENATE("IB USer Demographic as at ","-",TEXT(P{target_row},"dd-mmm-yyyy"))'
            )

        # Save the updated target workbook
        target_wb.save(target_file)
        print("Rows appended, report name added to column B, source file path added to column C, value from B2 added to column D, and data from Columns 1-7 added to Columns E-K with formula in Column F.")
    
    except FileNotFoundError:
        print(f"File not found: {source_file} or {target_file}")
    except KeyError:
        print(f"Sheet '{source_sheet_name}' or '{target_sheet_name}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Specify the file paths, sheet names, and row/column details
source_file = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Report Template.xlsx"
target_file = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx"
source_sheet_name = "Rpt-Maintain"  # This can be kept as a reference, but will not be used for the first sheet anymore
target_sheet_name = "Sheet1"         # Replace with the correct target sheet name
source_row_start_index = 0           # Start at the first row (0-based index) in the source sheet
target_row_start = 3                 # Starting row in the target sheet (1-based index)
target_column = 4                    # Starting column to insert into in the target sheet (1-based index)

# Run the function
extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_row_start, target_column)
#BLABLA