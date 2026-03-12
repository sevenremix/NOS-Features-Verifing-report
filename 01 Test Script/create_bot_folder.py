from feishu_uploader import FeishuUploader
import logging
import json

logging.basicConfig(level=logging.INFO)

print("==================================================")
print("Creating a Dedicated Folder for the Bot")
print("==================================================")

uploader = FeishuUploader()
uploader.user_access_token = None
uploader.refresh_token = None

if not uploader.authenticate():
    print("[ERROR] Authentication failed.")
else:
    print("[SUCCESS] Bot Authenticated.")
    url = f"{uploader.base_url}/drive/v1/files/create_folder"
    payload = {
        "name": "NOS_Automated_Reports_Bot_Owned",
        "folder_token": "" # Empty means root directory
    }
    
    result = uploader._call_api("POST", url, json=payload)
    
    if result and result.get("code") == 0:
        new_folder_token = result.get("data", {}).get("token")
        folder_url = result.get("data", {}).get("url")
        print("\n[SUCCESS] Dedicated folder created successfully!")
        print(f"Folder Name: NOS_Automated_Reports_Bot_Owned")
        print(f"[NEW FOLDER TOKEN]: {new_folder_token}")
        print(f"Folder URL: {folder_url}")
        print("\n[NEXT STEPS]:")
        print("1. Update 'parent_node' in feishu_secrets.json to this new token.")
        print("2. The Bot will automatically have full Read/Write permissions here.")
    else:
        print(f"\n[ERROR] Failed to create folder (Code {result.get('code', 'unknown')})")
        print(f"Message: {result.get('msg', 'No message')}")
