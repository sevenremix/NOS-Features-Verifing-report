from feishu_uploader import upload_to_feishu
import os

def test_upload():
    source_file = "Feature_TestCase_20260310.xlsx"
    target_name = "Feature_TestCase_20260310_test.xlsx"
    
    print(f"Starting test upload: {source_file} -> {target_name}")
    
    if not os.path.exists(source_file):
        print(f"Error: Source file '{source_file}' not found.")
        return

    # Call the uploader with a custom remote filename
    result = upload_to_feishu(source_file, remote_filename=target_name)
    
    if result and result.get("code") == 0:
        print("Success: Upload successful!")
        print(f"File Token: {result.get('data', {}).get('file_token')}")
    else:
        print("Failure: Upload failed.")
        if result:
            print(f"Error details: {result.get('msg')}")

if __name__ == "__main__":
    test_upload()
