
from feishu_uploader import FeishuUploader
import json

def find_spreadsheet_token():
    uploader = FeishuUploader()
    # Find the IPRAN NOS folder first
    folder_token, folder_type = uploader.find_file_by_name("IPRAN NOS")
    
    if not folder_token:
        print("Error: 'IPRAN NOS' folder not found.")
        return

    print(f"Listing files in 'IPRAN NOS' (Token: {folder_token})...")
    result = uploader.list_folder_files(folder_token)
    
    if result and result.get("code") == 0:
        files = result.get("data", {}).get("files", [])
        print(f"\n{'Name':<40} | {'Type':<10} | {'Token'}")
        print("-" * 80)
        for f in files:
            print(f"{f.get('name'):<40} | {f.get('type'):<10} | {f.get('token')}")
    else:
        print(f"Error listing files: {result}")

if __name__ == "__main__":
    find_spreadsheet_token()
