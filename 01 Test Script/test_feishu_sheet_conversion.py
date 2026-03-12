from feishu_uploader import upload_to_feishu
import os

def test_sheet_conversion():
    source_file = "Feature_TestCase_20260310.xlsx"
    target_name = "Feature_TestCase_Online_Sheet"
    
    print(f"Starting test: Upload and Convert to Sheet: {source_file}")
    
    if not os.path.exists(source_file):
        print(f"Error: Source file '{source_file}' not found.")
        return

    # Call the uploader with convert_to_sheet=True
    result = upload_to_feishu(
        source_file, 
        remote_filename=target_name, 
        convert_to_sheet=True
    )
    
    if result and result.get("code") == 0:
        print("Success: Upload and conversion to Sheet successful!")
        data = result.get("data", {})
        # Different API versions return token in different structures
        sheet_token = data.get("result", {}).get("token") or data.get("token")
        print(f"Online Sheet Token: {sheet_token}")
        print(f"You can now view it in your target folder.")
    else:
        print("Failure: Upload or conversion failed.")
        if result:
            print(f"Error details: {result.get('msg')}")

if __name__ == "__main__":
    test_sheet_conversion()
