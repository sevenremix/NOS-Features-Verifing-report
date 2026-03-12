from feishu_uploader import FeishuUploader
import json

def list_drive_bitables():
    uploader = FeishuUploader()
    parent_token = uploader.default_parent_node
    print(f"Searching for Bitables in parent node: {parent_token}...")
    
    result = uploader.list_folder_files(parent_token)
    if result and result.get("code") == 0:
        files = result.get("data", {}).get("files", [])
        bitables = [f for f in files if f.get("type") == "bitable"]
        if bitables:
            print("\nFound Bitables:")
            for b in bitables:
                print(f"Name: {b.get('name')}, Token: {b.get('token')}")
        else:
            print("No Bitables found in direct root. Searching all visible files...")
            # If not in root, we might need to search or list root specifically if parent_node is a subfolder
            # Try listing root (optional)
            
    else:
        print(f"Error: {result}")

if __name__ == "__main__":
    list_drive_bitables()
