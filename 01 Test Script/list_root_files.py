from feishu_uploader import FeishuUploader
import json

def list_root_files():
    uploader = FeishuUploader()
    parent_token = uploader.default_parent_node
    
    print(f"Listing files in parent node: {parent_token}...")
    result = uploader.list_folder_files(parent_token)
    
    if result and result.get("code") == 0:
        files = result.get("data", {}).get("files", [])
        print(f"\n{'Name':<40} | {'Type':<10} | {'Token'}")
        print("-" * 80)
        for f in files:
            print(f"{f.get('name'):<40} | {f.get('type'):<10} | {f.get('token')}")
    else:
        print(f"Error listing files: {result}")

if __name__ == "__main__":
    list_root_files()
