from feishu_uploader import FeishuUploader
uploader = FeishuUploader()
print("user_access_token starts with:", getattr(uploader, 'user_access_token', '')[:10])
res = uploader.get_sheet_values("PGwJsSGGqhdukAtkIH8curLWnnb", "0Atghg!A3:J3")
print("get_sheet_values result:", res)

