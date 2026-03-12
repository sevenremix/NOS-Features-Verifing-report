import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import os

def read_xlsx_robust(file_path, sheet_names=None):
    """
    Reads an xlsx file by parsing its XML directly, ignoring all styling.
    """
    data = {}
    with zipfile.ZipFile(file_path, 'r') as z:
        # 1. Get shared strings
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = z.read('xl/sharedStrings.xml')
            root = ET.fromstring(ss_xml)
            for t in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                shared_strings.append(t.text)

        # 2. Get sheet names and their IDs
        wb_xml = z.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)
        sheet_map = {} # Name -> Id
        for sheet in wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            name = sheet.get('name')
            r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            sheet_map[name] = r_id

        # 3. Get relationship map to find filenames
        rels_xml = z.read('xl/_rels/workbook.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        rel_map = {} # Id -> Filename
        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            id = rel.get('Id')
            target = rel.get('Target')
            rel_map[id] = target

        if sheet_names is None:
            sheet_names = sheet_map.keys()

        for name in sheet_names:
            if name not in sheet_map:
                continue
            
            r_id = sheet_map[name]
            target = rel_map[r_id]
            # Target might be relative, e.g. "worksheets/sheet1.xml"
            sheet_path = f"xl/{target}" if not target.startswith('xl/') else target
            
            if sheet_path not in z.namelist():
                continue
            
            sheet_xml = z.read(sheet_path)
            s_root = ET.fromstring(sheet_xml)
            
            rows = []
            for row in s_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                row_data = {}
                for cell in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    r = cell.get('r') # e.g. "A1"
                    t = cell.get('t') # Type: 's' for shared string
                    v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if v_elem is not None:
                        val = v_elem.text
                        if t == 's':
                            val = shared_strings[int(val)]
                        
                        # Extract column index
                        col_idx = ""
                        for char in r:
                            if char.isalpha(): col_idx += char
                            else: break
                        row_data[col_idx] = val
                rows.append(row_data)
            
            if rows:
                df = pd.DataFrame(rows)
                # Sort columns A, B, C...
                sorted_cols = sorted(df.columns, key=lambda x: (len(x), x))
                df = df[sorted_cols]
                # Use first row as header and drop it
                new_header = df.iloc[0]
                df = df[1:]
                df.columns = new_header
                data[name] = df

    return data

if __name__ == "__main__":
    file_path = "UAR600D-10XA Base Function Test Report.xlsx"
    print("Reading sheets using XML...")
    sheets = read_xlsx_robust(file_path, ["基础功能", "BFD", "IGMP", "拓展测试"])
    for name, df in sheets.items():
        print(f"\n--- {name} ---")
        # Search for header
        header_idx = -1
        for i in range(min(50, len(df))):
            row_vals = [str(x) for x in df.iloc[i].values]
            if any("Feature" in v for v in row_vals if isinstance(v, str)):
                header_idx = i
                break
        
        if header_idx != -1:
            print(f"Header found at row {header_idx + 1}:")
            print(df.iloc[header_idx].values)
            # Re-header the dataframe
            new_df = df.iloc[header_idx+1:].copy()
            new_df.columns = df.iloc[header_idx].values
            print("Data sample:")
            print(new_df.head(5))
        else:
            print("Header not found in first 50 rows.")
