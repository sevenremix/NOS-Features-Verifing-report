from feishu_uploader import FeishuUploader
import json

def diagnose_spreadsheet():
    uploader = FeishuUploader()
    # Found in previous logs for 'Feature_TestCase_Online_Sheet'
    spreadsheet_token = "Bxx5sTf42hI15WtLph8cR7rdnlc" 
    
    print(f"Listing sheets for spreadsheet: {spreadsheet_token}")
    result = uploader.list_sheets(spreadsheet_token)
    
    if result and result.get("code") == 0:
        sheets = result.get("data", {}).get("sheets", [])
        print(f"Found {len(sheets)} sheets:")
        for s in sheets:
            print(f"- Title: {s.get('title')}, ID: {s.get('sheetId')}")
    else:
        print("Failed to list sheets.")
        print(result)

if __name__ == "__main__":
    diagnose_spreadsheet()
