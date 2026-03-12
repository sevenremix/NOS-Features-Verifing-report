from feishu_uploader import FeishuUploader
import logging

logging.basicConfig(level=logging.INFO)

print("==================================================")
print("Testing Bot Access to Specific Folder")
print("Folder Token: BmvrfIyrFl01z5dnRhocSImpnUf")
print("==================================================")

uploader = FeishuUploader()

# Force Bot Identity (simulating cloud environment)
uploader.user_access_token = None
uploader.refresh_token = None

if not uploader.authenticate():
    print("[ERROR] Authentication failed.")
else:
    print("[SUCCESS] Bot Authenticated.")
    folder_token = "BmvrfIyrFl01z5dnRhocSImpnUf"
    
    print(f"\n[System] Attempting to list files in folder {folder_token}...")
    result = uploader.list_folder_files(folder_token)
    
    if result and result.get("code") == 0:
        files = result.get("data", {}).get("files", [])
        print(f"[SUCCESS] Access granted! Found {len(files)} items in this folder.")
        for f in files[:5]:
            print(f" - {f.get('name')} ({f.get('type')})")
        if len(files) > 5:
            print(" - ...")
    else:
        print(f"\n[ERROR] Access Denied! (Code {result.get('code', 'unknown')})")
        print(f"Server Message: {result.get('msg', 'No message')}")
        print("\n[REQUIRED ACTION]:")
        print(f"1. Open the folder in your browser: https://datrokeshu1.feishu.cn/drive/folder/{folder_token}")
        print("2. Click 'Share' (分享) -> 'Collaborators' (协作者).")
        print("3. Search for your Bot name and add it with 'Can Edit' or 'Can View' permissions.")
