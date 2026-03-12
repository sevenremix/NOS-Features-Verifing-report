import pandas as pd

file_path = r'UAR600D Feature Compare-1226.xlsx'
try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheet names: {xl.sheet_names}")
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet_name)
        print(f"Sheet [{sheet_name}] columns: {df.columns.tolist()}")
except Exception as e:
    print(f"Error: {e}")
