import pandas as pd

file_path = r'UAR600D Feature Compare-1226.xlsx'
sheet_name = 'NCS520 Software Features'
column_to_remove = 'UT NOS'

try:
    # Read the excel file
    xl = pd.ExcelFile(file_path)
    # Check if sheet exists
    if sheet_name not in xl.sheet_names:
        # If not, use the first sheet
        sheet_name = xl.sheet_names[0]
        print(f"Sheet '{sheet_name}' not found, using first sheet: {sheet_name}")

    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Check if the column exists
    if column_to_remove in df.columns:
        print(f"Removing column '{column_to_remove}' from {file_path}...")
        df.drop(columns=[column_to_remove], inplace=True)
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print("Column removed successfully.")
    else:
        print(f"Column '{column_to_remove}' not found in sheet '{sheet_name}'.")
        print(f"Columns found: {df.columns.tolist()}")

except Exception as e:
    print(f"Error: {e}")
