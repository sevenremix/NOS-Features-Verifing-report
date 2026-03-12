from feishu_uploader import FeishuUploader
import json

def get_bitable_views(app_token, table_id):
    uploader = FeishuUploader()
    url = f"{uploader.base_url}/bitable/v1/apps/{app_token}/tables/{table_id}/views"
    result = uploader._call_api("GET", url)
    if result and result.get("code") == 0:
        views = result.get("data", {}).get("items", [])
        print(f"Views in Table {table_id}:")
        for v in views:
            print(f"Name: {v.get('view_name')}, ViewID: {v.get('view_id')}")
    else:
        print(f"Error: {result}")

if __name__ == "__main__":
    APP_TOKEN = "D4CubiG74anXpkshxapcw13JnkA"
    TABLE_ID = "tbl0CHX04hJOaz1n"
    get_bitable_views(APP_TOKEN, TABLE_ID)
