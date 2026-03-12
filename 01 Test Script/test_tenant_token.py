from feishu_uploader import FeishuUploader
import logging

logging.basicConfig(level=logging.INFO)

print("==================================================")
print("Testing Tenant Access Token (Bot Identity)")
print("==================================================")

uploader = FeishuUploader()

# Simulate Cloud Environment by removing personal tokens
print("\n[System] Simulating Cloud Environment (Disabling Personal Tokens)...")
uploader.user_access_token = None
uploader.refresh_token = None

# Force authentication using only app_id and app_secret
print("\n[System] Authenticating as Tenant/Bot...")
success = uploader.authenticate()

if success:
    print("\n[SUCCESS] Authentication Successful!")
    token = getattr(uploader, 'tenant_access_token', '')
    if token:
        print(f"[LOCKED] Tenant Access Token generated successfully (starts with '{token[:15]}...')")
    
    print("\n[System] Testing API Access (Read Sheet)...")
    # Using the template token from the app
    test_token = "PGwJsSGGqhdukAtkIH8curLWnnb"
    test_range = "0Atghg"
    
    result = uploader.get_sheet_values(test_token, test_range)
    if result:
        print(f"[SUCCESS] Read Access Successful! Fetched {len(result)} records.")
    else:
        print("\n[ERROR] Read Access Failed/Forbidden!")
        print("[WARNING] ACTION REQUIRED: You must open the Feishu Sheet (PGwJsSGGqhdukAtkIH8curLWnnb) and add your App/Bot to the Collaborators list for this to work in the cloud without your personal token.")
else:
    print("\n[ERROR] Authentication Failed! Please check your app_id and app_secret.")
