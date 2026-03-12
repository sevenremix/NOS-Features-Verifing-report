from feishu_uploader import FeishuUploader
import json
import sys

# Set encoding to handle Chinese characters in windows console if needed
# sys.stdout.reconfigure(encoding='utf-8') 

def get_view_details(app_token, table_id, view_id):
    uploader = FeishuUploader()
    url = f"{uploader.base_url}/bitable/v1/apps/{app_token}/tables/{table_id}/views/{view_id}"
    result = uploader._call_api("GET", url)
    if result and result.get("code") == 0:
        view = result.get("data", {}).get("view", {})
        print(f"DEBUG: View Details: {json.dumps(view, ensure_ascii=False)}")
    else:
        print(f"Error fetching {view_id}: {result}")

if __name__ == "__main__":
    APP_TOKEN = "D4CubiG74anXpkshxapcw13JnkA"
    TABLE_ID = "tbl0CHX04hJOaz1n"
    # Try both IDs found earlier
    get_view_details(APP_TOKEN, TABLE_ID, "vewHn4PwrI")
    get_view_details(APP_TOKEN, TABLE_ID, "vewBynJP8r")
