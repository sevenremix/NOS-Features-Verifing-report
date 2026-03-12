from feishu_uploader import FeishuUploader
import os
import logging

logging.basicConfig(level=logging.INFO)

print("==================================================")
print("Verifying Bot WRITE Access to Specific Folder")
print("Target Folder: BmvrfIyrFl01z5dnRhocSImpnUf")
print("==================================================")

uploader = FeishuUploader()

# Force Bot Identity
uploader.user_access_token = None
uploader.refresh_token = None

if not uploader.authenticate():
    print("[ERROR] Authentication failed.")
else:
    print("[SUCCESS] Bot Authenticated.")
    folder_token = "BmvrfIyrFl01z5dnRhocSImpnUf"
    file_path = "dummy_test.txt"
    
    print(f"\n[System] Attempting to upload {file_path} to folder...")
    # Use upload_excel (it handles general files too) or _do_upload directly
    file_size = os.path.getsize(file_path)
    result = uploader._do_upload("Write_Test_Success.txt", file_path, folder_token, file_size)
    
    if result and result.get("code") == 0:
        file_token = result.get("data", {}).get("file_token")
        print(f"[SUCCESS] Write access confirmed! File uploaded with token: {file_token}")
        
        print("\n[System] Cleaning up (Deleting test file from Feishu)...")
        if uploader.delete_file(file_token):
            print("[SUCCESS] Cleanup finished.")
        else:
            print("[WARNING] Cleanup failed, please delete 'Write_Test_Success.txt' manually.")
    else:
        print(f"\n[ERROR] Write Access DENIED! (Code {result.get('code', 'unknown')})")
        print(f"Error Message: {result.get('msg', 'No message')}")
        print("\n[REQUIRED ACTION]:")
        print(f"1. Open the folder: https://datrokeshu1.feishu.cn/drive/folder/{folder_token}")
        print("2. In 'Share' (分享) settings, ensure your Bot is added as a 'Member' or 'Collaborator' with 'Can Edit' (可编辑) permissions.")
