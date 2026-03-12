# Feishu Uploader 使用指南

`feishu_uploader.py` 是一个功能强大的飞书自动化工具库，支持文件上传、格式转换、多身份验证以及在线表格的深度操作。

---

## 1. 核心类：`FeishuUploader`

`FeishuUploader` 类封装了所有与飞书开放平台交互的逻辑。

### 身份验证与管理
- **`__init__(secrets_path)`**: 初始化类实例，默认读取 `feishu_secrets.json`。
- **`authenticate()`**: 自动选择认证方式。优先使用**个人身份 (User Access Token)**，若无则降级为**机器人身份 (Tenant Access Token)**。
- **`refresh_user_token()`**: 核心自动化能力。当个人 Token 过期时，自动使用 `refresh_token` 获取新令牌并持久化到 JSON 文件。

### 文件操作 (Drive API)
- **`upload_excel(file_path, parent_node, remote_filename, convert_to_sheet)`**: 
    - **上传并转换**：如果 `convert_to_sheet=True`，上传完成后会自动转换为在线表格并删除原始 `.xlsx` 文件。
- **`delete_file(file_token, file_type)`**: 根据 Token 删除云端文件。
- **`list_folder_files(folder_token)`**: 列出文件夹下的所有文件（用于清理或查找）。
- **`find_file_by_name(name, folder_token)`**: 在指定目录下按名称查找文件 Token。

### 在线表格操作 (Sheets API)
- **`list_sheets(spreadsheet_token)`**: 获取文档内所有子页签的名称和 ID。
- **`add_new_sheet(spreadsheet_token, title)`**: 在文档中新建一个页签。
- **`update_sheet_values(spreadsheet_token, sheet_id, values)`**: 将二维列表数据写入指定页签。支持自动坐标计算（如 `A1:D10`）。

---

## 2. 便捷辅助函数

除了类方法外，脚本还提供了两个直接调用的函数：

- **`read_local_excel_first_sheet(file_path)`**: 使用 `openpyxl` 读取本地 Excel 首页并转换为二维列表，自动处理空值。
- **`upload_to_feishu(...)`**: 顶层封装接口，只需一行代码即可完成文件上传与转换。

---

## 3. 代码示例

### 案例 A：一键上传并转为在线表格
```python
from feishu_uploader import upload_to_feishu

# 上传并自动转换成飞书文档，原始 .xlsx 会被自动清理
upload_to_feishu("report.xlsx", convert_to_sheet=True)
```

### 案例 B：向已有表格追加历史页签
```python
from feishu_uploader import FeishuUploader, read_local_excel_first_sheet
import datetime

uploader = FeishuUploader()

# 1. 找到目标表格
token, _ = uploader.find_file_by_name("My_Target_Sheet")

# 2. 新建日期页签
new_title = datetime.datetime.now().strftime("%Y%m%d")
sheet_id = uploader.add_new_sheet(token, new_title)

# 3. 写入内容
data = read_local_excel_first_sheet("local_data.xlsx")
uploader.update_sheet_values(token, sheet_id, data)
```

---

## 4. 依赖项
- `requests`: 网络请求
- `openpyxl`: 处理本地 Excel
- `feishu_secrets.json`: 配置文件
