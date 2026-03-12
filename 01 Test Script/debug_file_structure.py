from feishu_uploader import FeishuUploader
import json

def check_file_structure():
    uploader = FeishuUploader()
    parent_token = uploader.default_parent_node
    result = uploader.list_folder_files(parent_token)
    if result and result.get("code") == 0:
        files = result.get("data", {}).get("files", [])
        if files:
            print(json.dumps(files[0], indent=2, ensure_ascii=False))
        else:
            print("No files found.")
    else:
        print(f"Error: {result}")

if __name__ == "__main__":
    check_file_structure()
