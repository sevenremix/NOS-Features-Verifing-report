from feishu_uploader import FeishuUploader
import json

def debug_folder_contents():
    uploader = FeishuUploader()
    folder_token = "BmvrfIyrFl01z5dnRhocSImpnUf"
    
    print(f"Listing contents of folder: {folder_token}")
    result = uploader.list_folder_files(folder_token)
    
    if result and result.get("code") == 0:
        files = result.get("data", {}).get("files", [])
        print(f"Found {len(files)} files/folders.")
        for f in files:
            print(f"- Name: {f.get('name')}, Type: {f.get('type')}, Token: {f.get('token')}")
    else:
        print("Failed to list folder contents.")
        print(result)

if __name__ == "__main__":
    debug_folder_contents()
