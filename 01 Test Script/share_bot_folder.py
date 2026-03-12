from feishu_uploader import FeishuUploader
import logging

logging.basicConfig(level=logging.INFO)

print("==================================================")
print("Sharing Bot-Owned Folder with User")
print("==================================================")

uploader = FeishuUploader()

# We need the user's Open ID to invite them as a collaborator.
# The easiest way to get the user's ID is to use the user_access_token (which identifies the user).
# We can use an endpoint like "Identify Current User" or "Get User Info".
# Another approach is to get the root folder info of the user, but since the user token is active, we can use it.

def get_user_openid():
    # Use the personal user token to find out "Who am I" (Open ID)
    url = f"{uploader.base_url}/authen/v1/user_info"
    if not getattr(uploader, 'user_access_token', None):
        print("[ERROR] No user_access_token found. Cannot identify user.")
        return None
        
    headers = {"Authorization": f"Bearer {uploader.user_access_token}"}
    try:
        import requests
        res = requests.get(url, headers=headers)
        data = res.json()
        if data.get("code") == 0:
            open_id = data.get("data", {}).get("open_id")
            print(f"[SUCCESS] Identified your Feishu Open ID: {open_id}")
            return open_id
        else:
            print(f"[ERROR] Failed to get user info: {data.get('msg')}")
            return None
    except Exception as e:
        print(f"[ERROR] Exception finding user info: {e}")
        return None

def share_folder(folder_token, open_id):
    # Now switch to Bot Identity to grant permissions on its folder
    uploader.user_access_token = None
    uploader.refresh_token = None
    if not uploader.authenticate():
        print("[ERROR] Bot Authentication failed.")
        return False
        
    # Drive API to add permission
    url = f"{uploader.base_url}/drive/v1/permissions/{folder_token}/members"
    payload = {
        "member_type": "openid",
        "member_id": open_id,
        "perm": "edit"
    }
    # Parameters for the permission request
    params = {"type": "folder"}
    
    result = uploader._call_api("POST", url, params=params, json=payload)
    if result and result.get("code") == 0:
        print("\n[SUCCESS] Folder successfully shared with you as an Editor!")
        return True
    else:
        print(f"\n[ERROR] Failed to share folder (Code {result.get('code')}): {result.get('msg')}")
        return False

# Execution Flow
open_id = get_user_openid()
if open_id:
    # This is the folder created in the previous step
    target_folder = "QarWfaJ6Cl84Y8dY0i6cSKQwnjf"
    print(f"\n[System] Asking Bot to grant 'Edit' permissions on folder '{target_folder}' to Open ID: {open_id}...")
    share_folder(target_folder, open_id)
