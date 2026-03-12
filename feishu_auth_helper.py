import requests
import json
import os
import webbrowser

def get_auth_tokens():
    secrets_path = "feishu_secrets.json"
    
    if not os.path.exists(secrets_path):
        print(f"Error: {secrets_path} not found.")
        return

    with open(secrets_path, 'r', encoding='utf-8') as f:
        secrets = json.load(f)
        app_id = secrets.get("app_id")
        app_secret = secrets.get("app_secret")

    if not app_id:
        print("Error: app_id is missing in secrets.json")
        return

    # 1. Generate Auth URL
    # Required scopes for Drive operations
    scopes = "drive:drive drive:file drive:file:upload"
    redirect_uri = "https://open.feishu.cn" # Default placeholder
    
    auth_url = f"https://open.feishu.cn/open-apis/authen/v1/index?app_id={app_id}&redirect_uri={redirect_uri}&scope={scopes}&state=mystate"
    
    print("\n" + "="*50)
    print("FEISHU USER AUTHORIZATION HELPER")
    print("="*50)
    print("\nStep 1: Open the following URL in your browser and authorize the app:")
    print(f"\n{auth_url}\n")
    
    # Try to open browser automatically
    webbrowser.open(auth_url)
    
    print("Step 2: After authorization, the browser will redirect to a URL like:")
    print("https://open.feishu.cn/?code=【这里的一串字符】&state=mystate")
    print("\nStep 3: Copy the 'code' value from the URL and paste it below:")
    
    auth_code = input("\nEnter code: ").strip()
    
    if not auth_code:
        print("Error: Code cannot be empty.")
        return

    # 2. Exchange code for tokens
    url = "https://open.feishu.cn/open-apis/authen/v1/access_token"
    payload = {
        "grant_type": "authorization_code",
        "code": auth_code
    }
    
    # Get app_access_token first
    app_token_url = "https://open.feishu.cn/open-apis/auth/v3/app_access_token/internal"
    app_res = requests.post(app_token_url, json={"app_id": app_id, "app_secret": app_secret})
    app_access_token = app_res.json().get("app_access_token")

    headers = {
        "Authorization": f"Bearer {app_access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload)
        print(f"Response status: {response.status_code}")
        
        try:
            data = response.json()
        except Exception:
            print(f"Error: Response is not JSON. Content: {response.text}")
            return

        if data.get("code") == 0:
            user_access_token = data.get("data", {}).get("access_token")
            refresh_token = data.get("data", {}).get("refresh_token")
            
            # Save back to secrets.json
            secrets["user_access_token"] = user_access_token
            secrets["refresh_token"] = refresh_token
            
            with open(secrets_path, 'w', encoding='utf-8') as f:
                json.dump(secrets, f, indent=4, ensure_ascii=False)
            
            print("\n" + "="*50)
            print("SUCCESS: Tokens acquired and saved to feishu_secrets.json")
            print("="*50)
        else:
            print(f"\nError: {data.get('msg')}")
            print(f"Full response: {data}")
            
    except Exception as e:
        print(f"\nException details: {str(e)}")

if __name__ == "__main__":
    get_auth_tokens()
