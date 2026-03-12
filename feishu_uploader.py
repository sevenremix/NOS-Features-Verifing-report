import requests
import json
import os
import logging
import time

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def col_index_to_name(n):
    """Convert a 1-based column index to an Excel column name (e.g., 1 -> A, 27 -> AA)."""
    name = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        name = chr(65 + remainder) + name
    return name

class FeishuUploader:
    """
    Dedicated class for uploading files to Feishu Drive and converting them to online Sheets.
    Supports both User Access Token (Personal Identity) and Tenant Access Token (Bot Identity).
    """
    def __init__(self, secrets_file="feishu_secrets.json"):
        self.secrets_file = secrets_file
        self.app_id = None
        self.app_secret = None
        self.default_parent_node = ""
        self.user_access_token = None
        self.refresh_token = None
        self.tenant_access_token = None
        self.token_expiry = 0
        self.base_url = "https://open.feishu.cn/open-apis"
        
        # Load secrets from file or streamlit
        self._load_secrets()

    def _load_secrets(self):
        """Load credentials from the secrets file or Streamlit secrets."""
        # Priority 1: Local JSON file
        if os.path.exists(self.secrets_file):
            try:
                with open(self.secrets_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.app_id = data.get('app_id')
                    self.app_secret = data.get('app_secret')
                    self.default_parent_node = data.get('parent_node')
                    self.user_access_token = data.get('user_access_token')
                    self.refresh_token = data.get('refresh_token')
                    logger.info(f"Feishu: Loaded secrets from {self.secrets_file}")
                    return
            except Exception as e:
                logger.error(f"Error loading {self.secrets_file}: {e}")

        # Priority 2: Streamlit Secrets
        try:
            import streamlit as st
            if "feishu" in st.secrets:
                s = st.secrets["feishu"]
                self.app_id = s.get("app_id")
                self.app_secret = s.get("app_secret")
                self.default_parent_node = s.get("parent_node")
                self.user_access_token = s.get("user_access_token")
                self.refresh_token = s.get("refresh_token")
                logger.info("Feishu: Loaded secrets from Streamlit `st.secrets`")
                return
        except ImportError:
            pass # Streamlit not installed
        except Exception as e:
            logger.error(f"Error loading secrets from Streamlit: {e}")

        logger.warning("Feishu: No secrets found in local file or Streamlit secrets.")


    def _save_tokens(self):
        """Save updated tokens back to the secrets file."""
        if not hasattr(self, 'secrets_file'):
            return
            
        try:
            with open(self.secrets_file, 'r', encoding='utf-8') as f:
                secrets = json.load(f)
            
            secrets["user_access_token"] = self.user_access_token
            secrets["refresh_token"] = self.refresh_token
            
            with open(self.secrets_file, 'w', encoding='utf-8') as f:
                json.dump(secrets, f, indent=4, ensure_ascii=False)
            logger.info("Feishu: Tokens updated and saved to secrets file.")

        except Exception as e:
            logger.error(f"Error saving tokens: {e}")

    def refresh_user_token(self):
        """Refresh user_access_token using refresh_token."""
        if not self.refresh_token:
            return False

        logger.info("Feishu: Attempting to refresh User Access Token...")
        url = f"{self.base_url}/authen/v1/refresh_access_token"
        
        # We need an app_access_token to call this
        app_token_url = f"{self.base_url}/auth/v3/app_access_token/internal"
        try:
            app_res = requests.post(app_token_url, json={"app_id": self.app_id, "app_secret": self.app_secret})
            app_access_token = app_res.json().get("app_access_token")
        except Exception as e:
            logger.error(f"Failed to get app_access_token for refresh: {e}")
            return False

        headers = {
            "Authorization": f"Bearer {app_access_token}",
            "Content-Type": "application/json; charset=utf-8"
        }
        payload = {
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token
        }
        
        try:
            response = requests.post(url, headers=headers, json=payload)
            data = response.json()
            if data.get("code") == 0:
                self.user_access_token = data.get("data", {}).get("access_token")
                self.refresh_token = data.get("data", {}).get("refresh_token")
                self._save_tokens()
                return True
            else:
                logger.error(f"Token refresh failed: {data.get('msg')}")
                logger.error("Help: Your refresh_token might have expired. Please run 'python feishu_auth_helper.py' to re-authorize.")
                return False
        except Exception as e:
            logger.error(f"Token refresh exception: {e}")
            return False

    def authenticate(self):
        """Determine authentication method and obtain token."""
        # Priority 1: User Identity (if tokens exist)
        if self.user_access_token:
            return True

        # Priority 2: Tenant Identity (Bot)
        if self.tenant_access_token and time.time() < (self.token_expiry - 60):
            return True

        if not self.app_id or not self.app_secret:
            logger.error("Authentication failed: Missing App ID or Secret.")
            return False

        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        payload = {"app_id": self.app_id, "app_secret": self.app_secret}
        
        try:
            response = requests.post(url, json=payload)
            response.raise_for_status()
            data = response.json()
            if data.get("code") == 0:
                self.tenant_access_token = data.get("tenant_access_token")
                expire_seconds = data.get("expire", 7200)
                self.token_expiry = time.time() + expire_seconds
                logger.info("Feishu: Authenticated as Tenant (Bot).")
                return True
            else:
                logger.error(f"Feishu Auth Failed: {data.get('msg')}")
                return False
        except Exception as e:
            logger.error(f"Feishu Auth Exception: {e}")
            return False

    def _call_api(self, method, url, headers=None, params=None, json=None, files=None, retry=True):
        """Centralized API caller with automatic token refresh and retry."""
        if headers is None:
            headers = {}
        
        # Ensure latest token is used
        token = self.user_access_token if self.user_access_token else self.tenant_access_token
        headers["Authorization"] = f"Bearer {token}"
        
        try:
            response = requests.request(method, url, headers=headers, params=params, json=json, files=files)
            
            # If response is not JSON, return it as is (or handle as error)
            if not response.text.strip().startswith('{'):
                return {"code": -1, "msg": f"Non-JSON response: {response.text[:100]}"}
                
            result = response.json()
            
            # Check for token expiration codes
            # 99991663, 99991677 (User token expired), 1061004 (Tenant token expired), 401 (Unauthorized), 99991668 (Invalid access token)
            if retry and result.get("code") in [99991663, 99991677, 99991668, 1061004, 401]:
                logger.info(f"API call failed (Code {result.get('code')}), attempting token refresh...")
                if self.user_access_token and self.refresh_user_token():
                    # Retry with new user token
                    return self._call_api(method, url, headers, params, json, files, retry=False)
                elif not self.user_access_token and self.authenticate():
                    # Retry with new tenant token
                    return self._call_api(method, url, headers, params, json, files, retry=False)
            
            return result
        except Exception as e:
            logger.error(f"API Request Exception ({method} {url}): {e}")
            return {"code": -1, "msg": str(e)}

    def upload_excel(self, file_path, parent_node=None, remote_filename=None, convert_to_sheet=False):
        """Upload an Excel file to Feishu Drive."""
        if not self.authenticate():
            return None
        
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return None

        original_ext = os.path.splitext(file_path)[1].lower()
        file_name = remote_filename if remote_filename else os.path.basename(file_path)
        if original_ext == ".xlsx" and not file_name.lower().endswith(".xlsx"):
            file_name += ".xlsx"

        target_node = parent_node if parent_node is not None else self.default_parent_node
        file_size = os.path.getsize(file_path)
        
        result = self._do_upload(file_name, file_path, target_node, file_size)
        
        if convert_to_sheet and result and result.get("code") == 0:
            file_token = result.get("data", {}).get("file_token")
            logger.info("File uploaded successfully. Initiating conversion to Sheet...")
            sheet_title = os.path.splitext(file_name)[0]
            import_result = self._create_import_task(file_token, sheet_title, target_node)
            
            if import_result and import_result.get("code") == 0:
                logger.info(f"Conversion finished. Deleting original source file {file_token}...")
                self.delete_file(file_token)
            return import_result
                
        return result

    def delete_file(self, file_token, file_type="file"):
        """Delete a file from Feishu Drive."""
        url = f"{self.base_url}/drive/v1/files/{file_token}"
        params = {"type": file_type}
        result = self._call_api("DELETE", url, params=params)
        if result.get("code") == 0:
            logger.info(f"File {file_token} deleted successfully.")
            return True
        logger.warning(f"Failed to delete file {file_token}: {result.get('msg')}")
        return False

    def _do_upload(self, file_name, file_path, target_node, file_size):
        url = f"{self.base_url}/drive/v1/files/upload_all"
        
        with open(file_path, 'rb') as f:
            file_content = f.read()
            
        files = {
            'file': (file_name, file_content),
            'file_name': (None, file_name),
            'parent_type': (None, "explorer"),
            'parent_node': (None, target_node),
            'size': (None, str(file_size))
        }
        logger.info(f"Uploading {file_name} to Feishu... (Retry-safe)")
        return self._call_api("POST", url, files=files)

    def _create_import_task(self, file_token, file_name, folder_token):
        """Create an asynchronous task to import/convert Excel to Sheet."""
        url = f"{self.base_url}/drive/v1/import_tasks"
        base_name = os.path.splitext(file_name)[0]
        payload = {
            "file_extension": "xlsx",
            "file_token": file_token,
            "type": "sheet",
            "file_name": base_name,
            "point": {"mount_type": 1, "mount_key": folder_token}
        }
        result = self._call_api("POST", url, json=payload)
        if result.get("code") == 0:
            ticket = result.get("data", {}).get("ticket")
            logger.info(f"Import task created. Ticket: {ticket}. Waiting for completion...")
            return self._wait_for_import_task(ticket)
        return result

    def _wait_for_import_task(self, ticket):
        """Poll the import task status until it completes."""
        url = f"{self.base_url}/drive/v1/import_tasks/{ticket}"
        max_retries = 30 # Increased to 30 (60 seconds) for larger files
        for i in range(max_retries):
            result = self._call_api("GET", url)
            if result.get("code") == 0:
                status = result.get("data", {}).get("result", {}).get("job_status")
                if status == 2:
                    final_token = result.get("data", {}).get("result", {}).get("token")
                    logger.info("Conversion successful!")
                    return result
                elif status == 3:
                    logger.error("Conversion failed on server.")
                    return result
            logger.info(f"Still converting... (attempt {i+1}/{max_retries})")
            time.sleep(2)
        logger.warning(f"Conversion polling timed out for ticket: {ticket}")
        return None

    def list_folder_files(self, folder_token):
        """List files in a specific folder."""
        url = f"{self.base_url}/drive/v1/files"
        params = {"folder_token": folder_token}
        return self._call_api("GET", url, params=params)

    def find_file_by_name(self, name, folder_token=None):
        """Find a file token by its name in a specific folder."""
        target_folder = folder_token if folder_token else self.default_parent_node
        result = self.list_folder_files(target_folder)
        if result and result.get("code") == 0:
            files = result.get("data", {}).get("files", [])
            for f in files:
                if f.get("name") == name:
                    return f.get("token"), f.get("type")
        return None, None

    def add_new_sheet(self, spreadsheet_token, title):
        """Add a new sheet to an existing spreadsheet (v2)."""
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/sheets_batch_update"
        payload = {
            "requests": [{"addSheet": {"properties": {"title": title}}}]
        }
        result = self._call_api("POST", url, json=payload)
        if result.get("code") == 0:
            replies = result.get("data", {}).get("replies", [])
            if replies:
                sheet_id = replies[0].get("addSheet", {}).get("properties", {}).get("sheetId")
                logger.info(f"New sheet '{title}' added (v2). ID: {sheet_id}")
                return sheet_id
        return None

    def list_sheets(self, spreadsheet_token):
        """List all sheets in a spreadsheet (v2)."""
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/metainfo"
        return self._call_api("GET", url)

    def update_sheet_values(self, spreadsheet_token, sheet_id, values):
        """Write values to a specific sheet (v2) using its internal sheetId."""
        if not values or not values[0]:
            return False
        rows, cols = len(values), len(values[0])
        c_name = col_index_to_name(cols)
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
        exact_range = f"{sheet_id}!A1:{c_name}{rows}"
        payload = {
            "valueRanges": [{"range": exact_range, "values": values}]
        }
        time.sleep(1)
        result = self._call_api("POST", url, json=payload)
        if result.get("code") == 0:
            logger.info("Successfully updated sheet values (v2).")
            return True
        logger.error(f"Failed to update values (v2, Code {result.get('code')}): {result.get('msg')}")
        return False

    def get_sheet_values(self, spreadsheet_token, sheet_id):
        """Read all values from a specific sheet (v2)."""
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!A1:Z5000"
        result = self._call_api("GET", url)
        if result.get("code") == 0:
            return result.get("data", {}).get("valueRange", {}).get("values", [])
        logger.error(f"Failed to get sheet values (Code {result.get('code')}): {result.get('msg')}")
        return None

    def get_sheet_id_by_name(self, spreadsheet_token, sheet_name):
        """Find the internal sheetId for a sheet given its display name."""
        result = self.list_sheets(spreadsheet_token)
        if result and result.get("code") == 0:
            sheets = result.get("data", {}).get("sheets", [])
            for s in sheets:
                if s.get("title") == sheet_name:
                    return s.get("sheetId")
        return None

def read_local_excel_first_sheet(file_path):
    """Read the first sheet of a local Excel file using openpyxl."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.worksheets[0]
        data = []
        for row in sheet.iter_rows(values_only=True):
            # Clean up None values to empty strings for Feishu
            data.append([str(cell) if cell is not None else "" for cell in row])
        return data
    except Exception as e:
        logger.error(f"Error reading local Excel: {e}")
        return None

def upload_to_feishu(file_path, parent_folder_token=None, remote_filename=None, convert_to_sheet=False):
    """
    Standard interface for external scripts to upload to Feishu.
    If convert_to_sheet is True, it will return the online Sheet's metadata.
    """
    uploader = FeishuUploader()
    return uploader.upload_excel(
        file_path, 
        parent_node=parent_folder_token, 
        remote_filename=remote_filename,
        convert_to_sheet=convert_to_sheet
    )

if __name__ == "__main__":
    print("Feishu Uploader initialized. Use upload_to_feishu(file_path, convert_to_sheet=True) to import as Sheets.")
