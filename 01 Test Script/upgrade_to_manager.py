import requests
import json
import logging
from feishu_uploader import FeishuUploader

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def upgrade_permission():
    uploader = FeishuUploader()
    target_folder = uploader.default_parent_node # This is "QarWfaJ6Cl84Y8dY0i6cSKQwnjf"
    
    if not target_folder:
        logger.error("No folder token found in secrets.")
        return

    # 1. Get User's Open ID using User Access Token
    logger.info("Step 1: Identifying user OpenID...")
    if not uploader.user_access_token:
        logger.error("Missing user_access_token. Please run auth helper if needed.")
        return
        
    auth_url = f"{uploader.base_url}/authen/v1/user_info"
    headers = {"Authorization": f"Bearer {uploader.user_access_token}"}
    try:
        res = requests.get(auth_url, headers=headers)
        user_data = res.json()
        if user_data.get("code") == 0:
            open_id = user_data.get("data", {}).get("open_id")
            logger.info(f"Detected OpenID: {open_id}")
        else:
            logger.error(f"Failed to identify user: {user_data.get('msg')}")
            return
    except Exception as e:
        logger.error(f"Error identifying user: {e}")
        return

    # 2. Authenticate as Bot to manage permissions
    logger.info("Step 2: Switching to Bot identity (Owner)...")
    # Clear user tokens to force Bot auth
    uploader.user_access_token = None
    uploader.refresh_token = None
    if not uploader.authenticate():
        logger.error("Bot authentication failed.")
        return

    # 3. Update permission level to 'full_access' (Manager)
    # Endpoint: PATCH /open-apis/drive/v1/permissions/:token/members/:member_id
    logger.info(f"Step 3: Upgrading OpenID {open_id} to 'full_access' (Manager) on folder {target_folder}...")
    
    update_url = f"{uploader.base_url}/drive/v1/permissions/{target_folder}/members/{open_id}"
    params = {
        "type": "folder",
        "member_type": "openid"
    }
    payload = {
        "perm": "full_access" # This corresponds to "Manager"
    }
    
    headers = {
        "Authorization": f"Bearer {uploader.tenant_access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    
    try:
        response = requests.patch(update_url, headers=headers, params=params, json=payload)
        result = response.json()
        if result.get("code") == 0:
            print("\n" + "="*50)
            print("🚀 权限升级成功！")
            print(f"您现在已成为目录 [{target_folder}] 的『管理 (Manager)』人。")
            print("请刷新飞书网页查看效果。")
            print("="*50)
        else:
            logger.error(f"API Error ({result.get('code')}): {result.get('msg')}")
            # Fallback: Try POST (Add) if PATCH fails for some reason
            logger.info("Retrying with POST (Add/Update) method...")
            add_url = f"{uploader.base_url}/drive/v1/permissions/{target_folder}/members"
            add_payload = {
                "member_type": "openid",
                "member_id": open_id,
                "perm": "full_access"
            }
            res_post = requests.post(add_url, headers=headers, params={"type": "folder"}, json=add_payload)
            data_post = res_post.json()
            if data_post.get("code") == 0:
                print("\n[SUCCESS] Permission upgraded via POST method.")
            else:
                logger.error(f"POST fallback also failed: {data_post.get('msg')}")
                
    except Exception as e:
        logger.error(f"Permission upgrade exception: {e}")

if __name__ == "__main__":
    upgrade_permission()
