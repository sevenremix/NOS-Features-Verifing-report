import zipfile
import xml.etree.ElementTree as ET
import pandas as pd

def read_xlsx_xml_robust(file_path, sheet_names=None):
    data = {}
    with zipfile.ZipFile(file_path, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = z.read('xl/sharedStrings.xml')
            root = ET.fromstring(ss_xml)
            for t in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                shared_strings.append(t.text)

        wb_xml = z.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)
        sheet_map = {}
        for sheet in wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            name = sheet.get('name')
            r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            sheet_map[name] = r_id

        rels_xml = z.read('xl/_rels/workbook.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        rel_map = {}
        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            id = rel.get('Id')
            target = rel.get('Target')
            rel_map[id] = target

        if sheet_names is None: sheet_names = sheet_map.keys()

        for name in sheet_names:
            if name not in sheet_map: continue
            r_id = sheet_map[name]
            target = rel_map[r_id]
            sheet_path = f"xl/{target}" if not target.startswith('xl/') else target
            if sheet_path not in z.namelist(): continue
            
            sheet_xml = z.read(sheet_path)
            s_root = ET.fromstring(sheet_xml)
            
            rows = []
            for row in s_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                row_data = {}
                for cell in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    r = cell.get('r')
                    t = cell.get('t')
                    v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if v_elem is not None:
                        val = v_elem.text
                        if t == 's': val = shared_strings[int(val)]
                        col_idx = "".join([c for c in r if c.isalpha()])
                        row_data[col_idx] = val
                rows.append(row_data)
            
            if rows:
                df = pd.DataFrame(rows)
                sorted_cols = sorted(df.columns, key=lambda x: (len(x), x))
                df = df[sorted_cols]
                header_idx = 0
                for i, r_dict in enumerate(rows):
                    if any(isinstance(v, str) and ("Feature" in v or "ID" in v) for v in r_dict.values()):
                        header_idx = i
                        break
                new_header = df.iloc[header_idx]
                df = df[header_idx+1:]
                df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(new_header)]
                data[name] = df.reset_index(drop=True)
    return data

TEST_FILE = "UAR600D-10XA Base Function Test Report.xlsx"
sheets = read_xlsx_xml_robust(TEST_FILE, ["基础功能", "BFD", "IGMP", "拓展测试"])

for name, df in sheets.items():
    print(f"\n--- Sheet: {name} ---")
    # Looking for features mentioned in the image
    targets = ["IPv4 and IPv6 unicast routing", "Virtual Routing and Forwarding (VRF)", "BGP4/BGP4+/MP-BGP"]
    
    # Try to find columns for Feature and Item
    feat_col = next((c for c in df.columns if "Feature" in c), None)
    item_col = next((c for c in df.columns if "Item" in c), None)
    
    if feat_col and item_col:
        for t in targets:
            matches = df[df[feat_col].str.contains(t, na=False, case=False)]
            if not matches.empty:
                print(f"Matches for '{t}':")
                print(matches[[item_col, feat_col]])
