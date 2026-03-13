import streamlit as st
import os
import time
import datetime
import json
from feishu_uploader import FeishuUploader, upload_to_feishu
from Feature_Case_Merging import run_merge

# --- CONFIGURATION LOADER ---
def get_config(key, default=None):
    # Try local file
    if os.path.exists('feishu_secrets.json'):
        try:
            with open('feishu_secrets.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
                if key in data: return data[key]
        except: pass
    # Try Streamlit Secrets
    try:
        import streamlit as st
        # Scoped
        if "feishu" in st.secrets and key in st.secrets["feishu"]:
            return st.secrets["feishu"][key]
        # Global
        if key in st.secrets:
            return st.secrets[key]
    except: pass
    return default

@st.cache_data(ttl=3600)
def get_dynamic_tokens():
    """Dynamically discover tokens if not provided in secrets."""
    uploader = FeishuUploader()
    if not uploader.authenticate():
        return None, None
    return uploader.discover_feature_source("Formatted_Feature_Source")

# --- INITIAL CONFIG ---
# Try to get from st.secrets or local file first
TEMPLATE_SS_TOKEN = get_config("template_ss_token")
TEMPLATE_SHEET_ID = get_config("template_sheet_id")

# If still missing, try dynamic discovery
if not TEMPLATE_SS_TOKEN or not TEMPLATE_SHEET_ID:
    d_token, d_sheet = get_dynamic_tokens()
    TEMPLATE_SS_TOKEN = TEMPLATE_SS_TOKEN or d_token or "PGwJsSGGqhdukAtkIH8curLWnnb"
    TEMPLATE_SHEET_ID = TEMPLATE_SHEET_ID or d_sheet or "0Atghg"

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="NOS Feature Merge Tool",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM CSS (Deep Grey Theme) ---
st.markdown("""
<style>
    /* Hide SideNav for Multi-page (since we are back to single page) */
    [data-testid="stSidebarNav"] { display: none !important; }
    
    /* Main background */
    .stApp, [data-testid="stAppViewContainer"] {
        background-color: #121212;
        color: #E0E0E0;
    }
    
    /* Sidebar */
    [data-testid="stSidebar"], .stSidebar {
        background-color: #1A1A1A !important;
        border-right: 1px solid #333333;
    }
    [data-testid="stSidebar"] * {
        color: #E0E0E0 !important;
    }
    
    /* Top Header/Toolbar */
    header[data-testid="stHeader"] {
        background-color: #121212 !important;
        color: #E0E0E0 !important;
    }
    
    /* Typography */
    h1, h2, h3, h4, .stMarkdown {
        color: #FFFFFF !important;
        font-family: 'Inter', sans-serif;
    }
    
    /* Buttons - Unified Style for all buttons */
    .stButton>button, .stLinkButton>a, .stFileUploader button {
        background-color: #1E1E1E !important; 
        color: #FFFFFF !important;
        border: 1px solid #333333 !important; 
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
        text-decoration: none !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        height: 45px !important;
        font-weight: 500 !important;
    }
    .stButton>button:hover, .stLinkButton>a:hover, .stFileUploader button:hover { 
        background-color: #2D2D2D !important; 
        border-color: #4D4D4D !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.5) !important;
    }
    
    /* Info/Code blocks/Notifications - Unified Dark Appearance */
    .stCode, div[data-testid="stCodeBlock"], .stNotification, div[data-testid="stNotification"], [data-testid="stSidebar"] pre {
        background-color: #1A1A1A !important;
        border: 1px solid #333333 !important;
        border-radius: 8px !important;
        color: #E0E0E0 !important;
    }
    /* Targeted fix for the white background in code blocks */
    [data-testid="stSidebar"] [data-testid="stCodeBlock"], [data-testid="stSidebar"] pre, [data-testid="stSidebar"] code {
        background-color: #1A1A1A !important;
        color: #FFFFFF !important;
    }
    /* Ensure code text is readable and doesn't have its own background */
    .stCode code, div[data-testid="stCodeBlock"] code {
        background-color: transparent !important;
        color: #FFFFFF !important;
    }

    /* File Uploader Section */
    .stFileUploader section { 
        background-color: #121212 !important; 
        border: 1px dashed #333333 !important; 
        border-radius: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- APP HEADER ---
st.title("🤖 NOS Feature & Case Processor")
st.markdown("集成化工具：**云端拉取规格** ➡️ **自动合并报表** ➡️ **同步飞书在线表**")
st.divider()

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ 系统状态")
    uploader = FeishuUploader()
    auth_success = uploader.authenticate()
    
    if auth_success:
        identity = "机器人" if uploader.tenant_access_token and not uploader.user_access_token else "个人身份"
        st.success(f"飞书连接：正常 ({identity})")
    else:
        st.error("飞书连接：未授权")
        if not uploader.app_id or not uploader.app_secret:
            st.warning("⚠️ 提示：未检测到 App ID 或 Secret。请在 Secrets 中配置 [feishu] 分组。")
        else:
            st.warning("⚠️ 提示：凭证无效或已过期，请检查权限设置。")
    
    st.divider()
    st.header("🔗 快捷入口")
    st.link_button("📂 打开飞书云端存储目录", 
                   "https://datrokeshu1.feishu.cn/drive/folder/QarWfaJ6Cl84Y8dY0i6cSKQwnjf", 
                   use_container_width=True)
    
    st.divider()
    st.markdown("### 📋 规格模板库")
    st.info(f"当前源：`Formatted_Feature_Source`")
    if st.button("查看模板详情", use_container_width=True):
        st.code(f"SS_TOKEN: {TEMPLATE_SS_TOKEN}\nSHEET_ID: {TEMPLATE_SHEET_ID}")
    
    st.divider()
    st.caption("v2.1.0 (Cloud Integrated) | Decoupled")

# --- MAIN UI ---
tab1, tab2, tab3 = st.tabs(["🚀 一键全自动处理", "📤 纯文件上传", "👁️ 云端数据预览"])

with tab1:
    st.subheader("1. 上传原始测试报告")
    test_file = st.file_uploader("点击或拖拽测试报告 (.xlsx)", type=["xlsx"], key="auto_merge")
    
    if test_file:
        st.markdown("---")
        st.subheader("2. 执行工作流")
        if st.button("✨ 开始全自动化处理", use_container_width=True):
            # temp storage
            temp_test_path = f"temp_report_{test_file.name}"
            with open(temp_test_path, "wb") as f:
                f.write(test_file.getvalue())
            
            with st.status("正在启动引擎...", expanded=True) as status:
                # Step 1: Fetch Cloud Data
                status.write("📡 正在从飞书云端拉取最新的规格模板...")
                feature_data = uploader.get_sheet_values(TEMPLATE_SS_TOKEN, TEMPLATE_SHEET_ID)
                if not feature_data:
                    status.update(label="❌ 无法获取云端规格", state="error")
                    st.stop()
                
                # Step 2: Merge in memory
                status.write("🧠 正在进行逻辑合并与统计计算...")
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_name = f"Feature_TestCase_report_{timestamp}.xlsx"
                merge_result_path = run_merge(feature_data, temp_test_path, output_file=output_name)
                
                if not merge_result_path:
                    status.update(label="❌ 合并逻辑执行出错", state="error")
                    st.stop()
                
                # 4. Upload to Feishu and Convert
                status.update(label="📤 正在上传并转换为云端在线表...", state="running")
                feishu_res = upload_to_feishu(merge_result_path, convert_to_sheet=True)
                if feishu_res and feishu_res.get("code") == 0:
                    st.write("✅ 已同步至飞书 Drive")
                else:
                    st.error(f"❌ 飞书上传失败: {feishu_res.get('msg') if feishu_res else '未知错误'}")
                    st.stop()
                
                # 5. Sync to Bitable Project Tracker
                status.update(label="🔄 正在同步项目进度至多维表格 (Bitable)...", state="running")
                from bitable_project_tracker import sync_report_to_bitable
                sync_res = sync_report_to_bitable(merge_result_path)
                if sync_res:
                    st.write("✅ Bitable 进度更新成功")
                else:
                    st.warning("⚠️ Bitable 同步失败，请检查脚本日志")
                
                status.update(label="✨ 处理全部完成！已生成报表并更新进度。", state="complete")
                st.balloons()
                st.success(f"🎉 **大功告成！** 报表已保存至飞书，项目进度已同步。")
            
            # Cleanup
            if os.path.exists(temp_test_path): os.remove(temp_test_path)
            if os.path.exists(merge_result_path): os.remove(merge_result_path)

with tab2:
    st.subheader("直接上传 Excel 到飞书驱动器")
    simple_file = st.file_uploader("选择任意 .xlsx 文件", type=["xlsx"], key="simple_upload")
    if simple_file:
        if st.button("直接同步 (不进行合并)", use_container_width=True):
            temp_path = f"simple_{simple_file.name}"
            with open(temp_path, "wb") as f: f.write(simple_file.getvalue())
            
            with st.spinner("上传中..."):
                result = upload_to_feishu(temp_path, convert_to_sheet=True)
                if result and result.get("code") == 0:
                    st.success("上传并转换成功！")
                else:
                    st.error(f"失败: {result}")
            if os.path.exists(temp_path): os.remove(temp_path)

with tab3:
    st.subheader("🔍 实时云端预览")
    
    # --- Drive File List Section ---
    st.markdown("### 📂 飞书目录文件列表")
    
    # Auto-fetch logic
    uploader = FeishuUploader()
    parent_token = uploader.default_parent_node
    
    with st.spinner("正在获取最新文件列表..."):
        folder_res = uploader.list_folder_files(parent_token)
        if folder_res and folder_res.get("code") == 0:
            files = folder_res.get("data", {}).get("files", [])
            if files:
                import pandas as pd
                file_df = pd.DataFrame(files)
                
                # Filter to only show sheet types (online spreadsheets)
                if 'type' in file_df.columns:
                    file_df = file_df[file_df['type'] == 'sheet']
                
                if file_df.empty:
                    st.info("该目录下暂无在线表格（Sheet）文件。")
                else:
                    # Sort by name descending (since we have timestamps in names) 
                    if 'created_time' in file_df.columns:
                        file_df['created_time'] = pd.to_numeric(file_df['created_time'])
                        file_df = file_df.sort_values(by='created_time', ascending=False)
                    else:
                        file_df = file_df.sort_values(by='name', ascending=False)
                    
                    st.write("---")
                    for index, row in file_df.iterrows():
                        col_info, col_btn = st.columns([3, 1])
                        with col_info:
                            create_time = "未知时间"
                            if 'created_time' in row:
                                try:
                                    create_time = datetime.datetime.fromtimestamp(int(row['created_time'])/1000).strftime('%Y-%m-%d %H:%M:%S')
                                except: pass
                            st.markdown(f"**{row['name']}**  \n`{row['type']}` | {create_time}")
                        with col_btn:
                            # Construct link directly to Feishu
                            feishu_url = f"https://feishu.cn/sheets/{row['token']}"
                            st.markdown(
                                f"""<a href="{feishu_url}" target="_blank" style="text-decoration: none;">
                                    <button style="
                                        width: 100%;
                                        background-color: #2D2D2D;
                                        color: white;
                                        padding: 8px 16px;
                                        border: 1px solid #00C853;
                                        border-radius: 8px;
                                        cursor: pointer;
                                        font-size: 14px;
                                        font-weight: bold;
                                        transition: all 0.3s;
                                    ">🌐 在飞书打开</button>
                                </a>""", 
                                unsafe_allow_html=True
                            )
                        st.divider()
            else:
                st.info("该目录下暂无文件。")
        else:
            st.error(f"无法获取列表: {folder_res.get('msg') if folder_res else '连接失败'}")
    
    st.divider()
    
    # --- Sheet Data Preview Section ---
    st.markdown(f"### 📊 规格数据预览 (`Formatted_Feature_Source`)")
    col1, col2 = st.columns(2)
    with col1:
        load_data = st.button("🔄 刷新并显示数据预览", use_container_width=True)
    with col2:
        view_original = st.toggle("🌐 显示飞书原始页面 (Iframe)", value=False)
    
    st.divider()
    
    if load_data:
        with st.spinner("正在从拉取云端数据..."):
            data = uploader.get_sheet_values(TEMPLATE_SS_TOKEN, TEMPLATE_SHEET_ID)
            if data:
                import pandas as pd
                # Helper to handle duplicate or None headers
                def clean_headers(headers):
                    seen = {}
                    new_headers = []
                    for i, h in enumerate(headers):
                        h_val = str(h) if h is not None and str(h).strip() != "" else f"Column_{i}"
                        if h_val in seen:
                            seen[h_val] += 1
                            new_headers.append(f"{h_val}_{seen[h_val]}")
                        else:
                            seen[h_val] = 0
                            new_headers.append(h_val)
                    return new_headers

                headers = clean_headers(data[0]) if len(data) > 0 else None
                df = pd.DataFrame(data[1:], columns=headers)
                st.dataframe(df, use_container_width=True, height=500)
                st.success(f"已加载 {len(data)} 行数据。")
            else:
                st.error("无法读取数据，请检查 Token 或网络。")
                
    if view_original:
        sheet_url = f"https://feishu.cn/sheets/{TEMPLATE_SS_TOKEN}"
        st.markdown(f"🔗 [点击跳转到飞书打开]({sheet_url})")
        st.components.v1.iframe(sheet_url, height=600, scrolling=True)

# --- FOOTER ---
st.markdown("<br><br>", unsafe_allow_html=True)
st.divider()
st.caption("技术栈：Streamlit + Feishu Open API + Pandas + Openpyxl")
