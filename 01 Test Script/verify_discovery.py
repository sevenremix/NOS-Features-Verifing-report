from feishu_uploader import FeishuUploader
import logging

logging.basicConfig(level=logging.INFO)

def test_discovery():
    uploader = FeishuUploader()
    if not uploader.authenticate():
        print("Auth failed.")
        return
        
    print("\n🔍 Testing Dynamic Discovery...")
    token, sheet_id = uploader.discover_feature_source("Formatted_Feature_Source")
    
    if token and sheet_id:
        print(f"✅ Discovery Success!")
        print(f"Spreadsheet Token: {token}")
        print(f"First Sheet ID: {sheet_id}")
    else:
        print(f"❌ Discovery Failed.")
        if not token:
            print("Could not find file 'Formatted_Feature_Source' in Bot's accessible folders.")
        elif not sheet_id:
            print("Found file but could not retrieve sheet IDs.")

if __name__ == "__main__":
    test_discovery()
