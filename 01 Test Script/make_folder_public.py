from feishu_uploader import FeishuUploader
import logging

logging.basicConfig(level=logging.INFO)

print("==================================================")
print("Setting Folder Link Sharing to 'Anyone Can Read'")
print("==================================================")

uploader = FeishuUploader()
uploader.user_access_token = None
uploader.refresh_token = None

if not uploader.authenticate():
    print("[ERROR] Bot Authentication failed.")
else:
    folder_token = "QarWfaJ6Cl84Y8dY0i6cSKQwnjf"
    url = f"{uploader.base_url}/drive/v1/permissions/{folder_token}/public"
    
    # PATCH /drive/v1/permissions/:token/public requires this specific structure
    payload = {
        "external_access": True,
        "security_entity": "anyone_can_view",
        "share_entity": "anyone",
        "link_share_entity": "anyone_readable", 
        "invite_external": True
    }
    
    params = {"type": "folder"}
    
    print(f"\n[System] Directing Bot to update public permissions for folder '{folder_token}'...")
    result = uploader._call_api("PATCH", url, params=params, json=payload)
    
    if result and result.get("code") == 0:
        print("\n[SUCCESS] Folder link sharing is now ON! (Anyone with the link can read)")
    else:
        print(f"\n[ERROR] Failed to set public permissions (Code {result.get('code')}): {result.get('msg')}")
        
        # If external sharing is organizationally blocked, try tenant only
        print("\n[System] Attempting fallback to Tenant-Only sharing...")
        fallback_payload = {
            "external_access": False,
            "security_entity": "anyone_can_view",
            "share_entity": "tenant",
            "link_share_entity": "tenant_readable",
            "invite_external": False
        }
        fallback_res = uploader._call_api("PATCH", url, params=params, json=fallback_payload)
        
        if fallback_res and fallback_res.get("code") == 0:
            print("\n[SUCCESS] Link sharing is ON! (Anyone IN YOUR ORGANIZATION with the link can read)")
        else:
            print(f"\n[ERROR] All attempts failed (Code {fallback_res.get('code')}): {fallback_res.get('msg')}")
            print("\n[MANUAL FIX REQUIRED]: Your Feishu Organization IT Administrator has strictly forbidden Apps from changing link sharing settings via API. You must manually turn on Link Sharing in the web interface.")
