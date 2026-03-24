import pandas as pd
import datetime
import json
import os
from feishu_uploader import FeishuUploader
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- CONFIGURATION LOADER ---
def get_config(key, default=None, required=False):
    val = None
    # Try local file
    if os.path.exists('feishu_secrets.json'):
        try:
            with open('feishu_secrets.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
                if key in data: val = data[key]
        except: pass
    # Try Streamlit Secrets
    if val is None:
        try:
            import streamlit as st
            if "feishu" in st.secrets and key in st.secrets["feishu"]:
                val = st.secrets["feishu"][key]
            elif key in st.secrets:
                val = st.secrets[key]
        except: pass

    # Validation: treat empty/whitespace strings as None
    if isinstance(val, str) and not val.strip():
        val = None

    if val is not None: return val

    if required and default is None:
        import streamlit as st
        st.error(f"❌ 缺少必要配置项: `{key}` (Bitable)")
        st.info("请检查 `feishu_secrets.json` 或云端配置，并确保值不为空。")
        st.stop()
    return default

# --- BITABLE SETTINGS ---
# Strict validation for all Bitable related keys
BITABLE_APP_TOKEN = get_config("bitable_app_token", required=True)
BITABLE_TABLE_ID = get_config("bitable_table_id", required=True)
BITABLE_VIEW_ID = get_config("bitable_view_id", required=True)
BITABLE_FIELD_ID = get_config("bitable_field_id", "fld0tWfQiB") # Field ID is still optional with default

def update_view_filter(uploader, app_token, table_id, view_id, date_str):
    """
    Updates the filter of a specific view to match the given date string.
    """
    try:
        url = f"{uploader.base_url}/bitable/v1/apps/{app_token}/tables/{table_id}/views/{view_id}"
        
        # Payload based on Feishu API for updating view property
        payload = {
            "property": {
                "filter_info": {
                    "conjunction": "and",
                    "conditions": [
                        {
                            "field_id": BITABLE_FIELD_ID,
                            "operator": "contains",
                            "value": json.dumps([date_str]) # Feishu expects JSON string for list values in filters
                        }
                    ]
                }
            }
        }
        
        logger.info(f"Updating View {view_id} filter to '{date_str}'...")
        result = uploader._call_api("PATCH", url, json=payload)
        
        if result and result.get("code") == 0:
            logger.info("View filter updated successfully!")
            return True
        else:
            logger.error(f"Failed to update view filter: {result.get('msg')}")
            return False
    except Exception as e:
        logger.error(f"Error updating view filter: {e}")
        return False

def sync_report_to_bitable(excel_file_path):
    """
    1. Reads the footer (Total Statistics) of the generated Excel report.
    2. Appends a new row to the specified Feishu Bitable.
    3. Updates the '本周数据' view filter.
    """
    try:
        logger.info(f"Extracting verification numbers from {excel_file_path}...")
        df = pd.read_excel(excel_file_path)
        last_row = df.iloc[-1]
        
        def extract_num(val):
            val_str = str(val)
            if "Feature Pass Num" in val_str:
                return int(val_str.split("Feature Pass Num")[-1].strip())
            return 0

        feat_pass_num = extract_num(last_row.get("Test Result", ""))
        cli_pass_num = extract_num(last_row.get("CLI State", ""))
        
        logger.info(f"Extracted Values -> Feature Pass: {feat_pass_num}, CLI Pass: {cli_pass_num}")
        
        uploader = FeishuUploader()
        if not uploader.authenticate():
             logger.error("Feishu authentication failed.")
             return False
             
        now = datetime.datetime.now()
        date_simple = now.strftime("%Y/%m/%d")
        today_str = f"截至{date_simple}"
        
        # 1. Add Record
        record_url = f"{uploader.base_url}/bitable/v1/apps/{BITABLE_APP_TOKEN}/tables/{BITABLE_TABLE_ID}/records"
        payload = {
            "fields": {
                "Items": today_str,
                "Feature验证数": feat_pass_num,
                "CLI验证数": cli_pass_num
            }
        }
        
        logger.info(f"Appending to Bitable: {today_str}...")
        res_record = uploader._call_api("POST", record_url, json=payload)
        
        if res_record and res_record.get("code") == 0:
            logger.info("Record added successfully.")
            # 2. Update View Filter to show THIS date specifically
            update_view_filter(uploader, BITABLE_APP_TOKEN, BITABLE_TABLE_ID, BITABLE_VIEW_ID, date_simple)
            return True
        else:
            logger.error(f"Bitable record add failed: {res_record.get('msg')}")
            return False
            
    except Exception as e:
        logger.error(f"Error during Bitable sync: {e}")
        return False

if __name__ == "__main__":
    # Test with a dummy file if needed
    import glob
    reports = glob.glob("Feature_TestCase_report_*.xlsx")
    if reports:
        sync_report_to_bitable(sorted(reports)[-1])
