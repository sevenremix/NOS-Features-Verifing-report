import streamlit as st
import os
import time
import datetime
from feishu_uploader import FeishuUploader, upload_to_feishu

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="Feishu Cloud Uploader",
    page_icon="☁️",
    layout="centered"
)

# --- CUSTOM CSS (Deep Grey Theme) ---
st.markdown("""
<style>
    /* Main background */
    .stApp {
        background-color: #121212;
        color: #E0E0E0;
    }
    
    /* Header and Title */
    h1, h2, h3 {
        color: #FFFFFF !important;
        font-family: 'Inter', sans-serif;
        font-weight: 700;
    }
    
    /* Button Styling */
    .stButton>button {
        background-color: #2D2D2D;
        color: #FFFFFF;
        border: 1px solid #3D3D3D;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #3D3D3D;
        border-color: #4D4D4D;
        box-shadow: 0 4px 12px rgba(0,0,0,0.5);
    }
    
    /* File Uploader area */
    .stFileUploader section {
        background-color: #1E1E1E !important;
        border: 1px dashed #3D3D3D !important;
        border-radius: 12px !important;
    }
    
    /* Success/Info boxes */
    .stAlert {
        background-color: #1A1A1A !important;
        border-left-color: #00C853 !important;
        color: #E0E0E0 !important;
    }
    
    /* Sidebar */
    .stSidebar {
        background-color: #0F0F0F !important;
    }
</style>
""", unsafe_allow_html=True)

# --- APP HEADER ---
st.title("☁️ Feishu Cloud Excel Uploader")
st.markdown("快速上传本地 Excel 并自动转换为**飞书在线表格**。")
st.divider()

# --- SIDEBAR: Status & Settings ---
with st.sidebar:
    st.header("⚙️ 状态")
    uploader_core = FeishuUploader()
    if uploader_core.authenticate():
        st.success("身份验证：已连接 (User/Bot)")
    else:
        st.error("身份验证：未连接")
        st.info("请确保 feishu_secrets.json 配置正确")
    
    st.divider()
    st.caption("v1.0.0 | Powered by Feishu Open Platform")

# --- MAIN UI ---
uploaded_file = st.file_uploader("选择一个 Excel 文件 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.info(f"📁 已选择: {uploaded_file.name} ({uploaded_file.size / 1024:.1f} KB)")
    
    # Optional Rename
    custom_name = st.text_input("重命名 (可选)", placeholder="留空则使用原文件名")
    
    # Action Button
    if st.button("🚀 开始同步至飞书", use_container_width=True):
        try:
            # 1. Save to temp local file
            temp_path = os.path.join(os.getcwd(), f"temp_{uploaded_file.name}")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getvalue())
            
            # 2. Start Process with status
            with st.status("正在处理...", expanded=True) as status:
                status.write("正在建立安全连接...")
                time.sleep(0.5)
                
                status.write("正在上传文件到 Feishu Drive...")
                remote_name = custom_name if custom_name else uploaded_file.name
                
                # Call the high-level function
                result = upload_to_feishu(
                    temp_path, 
                    remote_filename=remote_name, 
                    convert_to_sheet=True
                )
                
                if result and result.get("code") == 0:
                    status.update(label="✅ 同步完成！", state="complete", expanded=False)
                    st.balloons()
                    st.success(f"**成功！** 文件 '{remote_name}' 已转为在线表格。")
                    st.info("飞书云端已自动清理临时 .xlsx 文件，仅保留在线表格。")
                else:
                    status.update(label="❌ 任务失败", state="error", expanded=True)
                    st.error(f"处理出错: {result.get('msg') if result else '未知错误'}")
            
            # 3. Local Cleanup
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
        except Exception as e:
            st.error(f"异常错误: {str(e)}")

# --- FOOTER ---
st.markdown("<br><br>", unsafe_allow_html=True)
st.divider()
st.caption("注意：转换过程由于飞书后端异步处理，可能需要 1-3 秒完成。")
