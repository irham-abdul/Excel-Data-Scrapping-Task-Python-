import pandas as pd #import library first(pip install)
from openpyxl import load_workbook
import os
#another comment izz
#another commet irham
print("try")
def extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_row_start, target_column):
    try:
        # Load source Excel file as a DataFrame (reading Sheet2 for the data)
        df_source = pd.read_excel(source_file, sheet_name=source_sheet_name, index_col=None)
        
        # Get name of the report from the source file without the extension
        report_name = os.path.basename(source_file).split('.')[0]  
        
        # Get the absolute path of the source file
        source_file_path = os.path.abspath(source_file)
        
        # Load the source workbook and reference the first sheet
        source_wb = load_workbook(source_file)
        source_ws = source_wb[source_wb.sheetnames[0]]  # Access the first sheet, base index 0
        
        # Extract value from cell B2 of the first sheet
        value_b2 = source_ws['B2'].value 
        print(f"Value from B2: {value_b2}")  
        
        # Load the target workbook and worksheet
        target_wb = load_workbook(target_file)
        target_ws = target_wb[target_sheet_name]
        
        # Insert the report name into column B 
        #all starting from row 3
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=2, value=report_name)  
        
        # Insert the source file path into column C 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=3, value=source_file_path)  
        
         # Insert the value from B2 of the first sheet into column D 
        for row in range(target_row_start, target_row_start + len(df_source)):
            target_ws.cell(row=row, column=4, value=value_b2)  
        
        # Loop through each row in the source (Sheet2), starting from the specified start index
        for row_offset in range(len(df_source)):
            row_data = df_source.iloc[row_offset, 1:8] 
            
            target_row = target_row_start + row_offset
            
            for col_offset, value in enumerate(row_data):
                target_ws.cell(row=target_row, column=5 + col_offset, value=value)  

            # Insert CONCATENATE formula into Column F for current row
            target_ws.cell(
                row=target_row, 
                column=6,  # Column F is column 6 (1-based index)
                value=f'=CONCATENATE("IB USer Demographic as at ","-",TEXT(P{target_row},"dd-mmm-yyyy"))'
            )
        # Save updated target workbook
        target_wb.save(target_file)
        print("Rows appended, report name added to column B, source file path added to column C, value from B2 added to column D, and data from Columns 1-7 added to Columns E-K with formula in Column F.")
    
    except FileNotFoundError:
        print(f"File not found: {source_file} or {target_file}")
    except KeyError:
        print(f"Sheet '{source_sheet_name}' or '{target_sheet_name}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# File paths, sheet names, and row/column details
source_file = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Report Template.xlsx"
target_file = r"C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx"
source_sheet_name = "Rpt-Maintain"  
target_sheet_name = "Sheet1"         
source_row_start_index = 0           # Start at the first row in the source sheet, Base 0
target_row_start = 3                 # Starting row in the target sheet, 1-based index
target_column = 4                    # Starting column to insert into in the target sheet, 1-based index

extract_and_append_rows(source_file, target_file, source_sheet_name, target_sheet_name, source_row_start_index, target_row_start, target_column)

#portion to write from excel file to text.

sheets_dict = pd.read_excel(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx', sheet_name=None)

df = pd.read_excel(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final Extraction.xlsx', sheet_name='Sheet1')   #source ecxel to convert to text
df.to_csv(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.txt', sep='\t', index=False)           #target output file
print("Succesfully output excel data to Final_Extraction.txt")
df.to_csv(r'C:\Users\mirham\Downloads\INTERN FILE\TASK 3\PART 2\Final_Extraction.csv', sep='\t', index=False)           #target output file in csv form
print("Succesfully output excel data to Final_Extraction.csv")
