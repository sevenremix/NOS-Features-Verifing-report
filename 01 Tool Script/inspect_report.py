import pandas as pd
from openpyxl import load_workbook

def inspect_report(file_path):
    print(f"Inspecting {file_path}...")
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        return

    print("\nHeaders:")
    print(df.columns.tolist())
    
    print("\nSummary Rows (Level 1/2) and Statistics:")
    # Using new headers: Feature（二级）, Test Result, CLI State
    try:
        print(df.head(20)[['Feature（二级）', 'Test Result', 'CLI State']])
    except KeyError as e:
        print(f"Column not found in head: {e}")

    # Look for the Total Statistics row at the end
    print("\nLast 5 rows (including Total Statistics):")
    # Using new headers: Feature（一级）, Case items, Test Result, CLI State
    try:
        print(df.tail(5)[['Feature（一级）', 'Case items', 'Test Result', 'CLI State']])
    except KeyError as e:
        print(f"Column not found in tail: {e}")

if __name__ == "__main__":
    inspect_report("Feature_TestCase.xlsx")
