from feishu_uploader import FeishuUploader, read_local_excel_first_sheet
import datetime
import os

def append_to_existing_sheet():
    uploader = FeishuUploader()
    local_file = "Feature_TestCase_20260310.xlsx"
    target_spreadsheet_name = "Feature_TestCase_Online_Sheet"
    
    print(f"1. Reading local file: {local_file}")
    if not os.path.exists(local_file):
        print("Error: Local file not found.")
        return
        
    data = read_local_excel_first_sheet(local_file)
    if not data:
        print("Error: Could not read data from local file.")
        return

    print(f"2. Finding existing spreadsheet: {target_spreadsheet_name}")
    spreadsheet_token, file_type = uploader.find_file_by_name(target_spreadsheet_name)
    
    if not spreadsheet_token:
        print(f"Target spreadsheet '{target_spreadsheet_name}' not found. Creating a new one instead...")
        # If not found, we create a new one to ensure the test can proceed
        result = uploader.upload_excel(local_file, remote_filename=target_spreadsheet_name, convert_to_sheet=True)
        if result and result.get("code") == 0:
            print("Successfully created a new online spreadsheet.")
            # Note: The newly created one already has the data, so we might stop here or add another dated sheet.
            # For this requirement, we'll try to find its token again if it wasn't returned clearly.
            # But let's assume if it's missing, we just created it and it's fine.
            return
        else:
            print("Failed to create spreadsheet.")
            return

    # 3. Add a new sheet with date+time
    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    print(f"3. Adding new sheet: {now_str}")
    sheet_id = uploader.add_new_sheet(spreadsheet_token, now_str)
    
    if not sheet_id:
        print("Failed to add new sheet.")
        return

    # 4. Write data to the new sheet
    print("4. Writing data to the new sheet...")
    if uploader.update_sheet_values(spreadsheet_token, sheet_id, data):
        print(f"Success! New sheet '{now_str}' added to '{target_spreadsheet_name}'.")
    else:
        print("Failed to write data.")

if __name__ == "__main__":
    append_to_existing_sheet()
