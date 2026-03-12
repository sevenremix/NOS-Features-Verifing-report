from feishu_uploader import FeishuUploader
import json

def get_bitable_tables(app_token):
    uploader = FeishuUploader()
    url = f"{uploader.base_url}/bitable/v1/apps/{app_token}/tables"
    result = uploader._call_api("GET", url)
    if result and result.get("code") == 0:
        tables = result.get("data", {}).get("items", [])
        print(f"Tables in Bitable {app_token}:")
        for t in tables:
            print(f"Name: {t.get('name')}, TableID: {t.get('table_id')}")
    else:
        print(f"Error: {result}")

if __name__ == "__main__":
    APP_TOKEN = "D4CubiG74anXpkshxapcw13JnkA"
    get_bitable_tables(APP_TOKEN)
