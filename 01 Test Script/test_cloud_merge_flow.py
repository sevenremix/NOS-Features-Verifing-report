from feishu_uploader import FeishuUploader, upload_to_feishu
from Feature_Case_Merging import run_merge
import glob
import os
import datetime

def orchestrate_cloud_merge():
    uploader = FeishuUploader()
    
    # Configuration
    TEMPLATE_SS_TOKEN = "PGwJsSGGqhdukAtkIH8curLWnnb"
    TEMPLATE_SHEET_ID = "0Atghg"
    
    print("--- Phase 1: Pulling Template Data from Feishu Cloud ---")
    raw_feature_data = uploader.get_sheet_values(TEMPLATE_SS_TOKEN, TEMPLATE_SHEET_ID)
    
    if not raw_feature_data:
        print("Error: Could not retrieve data from Feishu Sheet.")
        return

    print(f"Successfully pulled {len(raw_feature_data)} rows from cloud.")

    print("\n--- Phase 2: Identifying Local Test Report ---")
    test_files = glob.glob("UAR600D-10XA Base Function Test Report (*).xlsx")
    if not test_files:
        print("Error: No local test report found.")
        return
    
    test_file = sorted(test_files)[-1]
    print(f"Using latest test report: {test_file}")

    print("\n--- Phase 3: Running Decoupled Merging Logic ---")
    # We pass the raw data list directly to run_merge
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"Feature_TestCase_report_{timestamp}.xlsx"
    result_path = run_merge(raw_feature_data, test_file, output_file=output_filename)

    if not result_path or not os.path.exists(result_path):
        print("Error: Merging process failed.")
        return

    print(f"Merged report generated at: {result_path}")

    print("\n--- Phase 4: Syncing Result back to Feishu Cloud ---")
    upload_res = upload_to_feishu(result_path, convert_to_sheet=True)
    
    if upload_res and upload_res.get("code") == 0:
        print("\n[SUCCESS] Full Cloud-Decoupled workflow completed!")
        # Some Feishu API versions don't return the token in the final polling response
        # but the conversion is successful if job_status is 2.
        sheet_token = upload_res.get('data', {}).get('result', {}).get('token')
        if sheet_token:
            print(f"Resulting Sheet Token: {sheet_token}")
        else:
            print("Note: Conversion successful, but Sheet Token was not returned in the polling response.")
            print("You can find the new sheet in your Feishu 'IPRAN NOS' (or designated) folder.")
        
        # Cleanup local file (as simulated in cloud environment)
        # os.remove(result_path)
        # print(f"Cleaned up temporary file: {result_path}")
    else:
        print(f"\n[FAILED] Upload failed: {upload_res}")

if __name__ == "__main__":
    orchestrate_cloud_merge()
