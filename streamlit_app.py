import streamlit as st
import google.generativeai as genai
import os
import json
import re
import pandas as pd
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

st.set_page_config(page_title="供應商資料自動辨識與主檔系統", layout="wide")

st.title("📑 供應商資料自動辨識與主檔系統 (Gemini AI 最新版)")
st.caption("支援上傳格式：PDF (.pdf), Word (.docx / .doc), Excel (.xlsx / .xls)")

api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
drive_folder_id = st.secrets.get("drive_folder_id") or os.environ.get("DRIVE_FOLDER_ID")

if not api_key:
    st.warning("⚠️ 請先於 secrets.toml 中設定 GEMINI_API_KEY。")

genai.configure(api_key=api_key)

def get_drive_service():
    if "gcp_service_account" in st.secrets:
        creds_info = dict(st.secrets["gcp_service_account"])
        scopes = ["https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scopes)
        return build("drive", "v3", credentials=creds)
    return None

def clean_folder_id(folder_id):
    if not folder_id:
        return ""
    match = re.search(r'[a-zA-Z0-9_-]{25,50}', str(folder_id))
    if match:
        return match.group(0)
    return str(folder_id).strip()

cleaned_folder_id = clean_folder_id(drive_folder_id)

def update_vendor_excel_on_drive(vendor_data, folder_id):
    service = get_drive_service()
    if not service or not folder_id:
        return False, "未設定 GCP Service Account 或 drive_folder_id"
    
    file_name = "供應商資料主檔.xlsx"
    query = f"'{folder_id}' in parents and name='{file_name}' and trashed=false"
    
    try:
        results = service.files().list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        items = results.get("files", [])
    except Exception as e:
        return False, f"查詢雲端資料夾失敗 (Folder ID: {folder_id}): {e}"
    
    new_df = pd.DataFrame([vendor_data])
    
    try:
        if items:
            file_id = items[0]["id"]
            request = service.files().get_media(fileId=file_id)
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            
            existing_df = pd.read_excel(fh)
            updated_df = pd.concat([existing_df, new_df], ignore_index=True)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                updated_df.to_excel(writer, index=False)
            output.seek(0)
            
            media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
            service.files().update(
                fileId=file_id,
                media_body=media,
                supportsAllDrives=True
            ).execute()
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                new_df.to_excel(writer, index=False)
            output.seek(0)
            
            file_metadata = {
                "name": file_name,
                "parents": [folder_id]
            }
            media = MediaIoBaseUpload(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
            service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True
            ).execute()
            
        return True, "成功寫入雲端主檔"
    except Exception as e:
        return False, f"檔案寫入 Google Drive 失敗: {e}"

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 上傳檔案預覽與資訊")
    uploaded_file = st.file_uploader("請選擇要上傳的供應商文件 (PDF, Word, Excel)", type=["pdf", "docx", "doc", "xlsx", "xls"])
    
    if uploaded_file:
        st.info(f"📄 已選擇檔案：`{uploaded_file.name}`")
        st.success("✅ 原始檔案已載入準備辨識！")

with col2:
    st.subheader("👁️ AI 欄位辨識與核對編輯")
    if uploaded_file and api_key:
        file_key = f"parsed_scope_{uploaded_file.name}"
        
        if file_key not in st.session_state:
            with st.spinner("AI 正在辨識文件、服務範疇與匯款資訊..."):
                try:
                    model = genai.GenerativeModel("gemini-1.5-flash")
                    bytes_data = uploaded_file.getvalue()
                    
                    prompt = """
                    請從這份供應商文件中提取以下欄位資訊，並直接以純 JSON 格式輸出：
                    {
                      "vendor_name": "供應商名稱",
                      "tax_id": "統一編號",
                      "service_scope": "服務範疇 (例如：軟體開發、硬體設備、公關行銷、諮詢服務、零組件供應等主營業務/營業項目)",
                      "contact_person": "聯絡人",
                      "phone": "聯絡電話",
                      "email": "電子郵件",
                      "address": "地址",
                      "bank_name": "銀行名稱",
                      "branch_name": "分行名稱",
                      "bank_code": "銀行代碼 (包含分行代碼若有)",
                      "bank_account_name": "戶名",
                      "bank_account_number": "匯款帳號"
                    }
                    如欄位不存在請留空字串 ""。
                    """
                    
                    response = model.generate_content([
                        {"mime_type": uploaded_file.type or "application/pdf", "data": bytes_data},
                        prompt
                    ])
                    
                    raw_text = response.text.strip()
                    # 使用不需要跳脫的非貪婪匹配，徹底避開 
 造成的跨行 SyntaxError
                    cleaned_text = re.sub(r"^```[a-zA-Z]*", "", raw_text)
                    cleaned_text = re.sub(r"```$", "", cleaned_text).strip()
                    
                    parsed_json = json.loads(cleaned_text)
                    st.session_state[file_key] = parsed_json
                except Exception as e:
                    st.error(f"❌ 文件 AI 辨識發生錯誤: {e}")
                    st.session_state[file_key] = {
                        "vendor_name": "", "tax_id": "", "service_scope": "", "contact_person": "", "phone": "", "email": "", "address": "",
                        "bank_name": "", "branch_name": "", "bank_code": "", "bank_account_name": "", "bank_account_number": ""
                    }

        if file_key in st.session_state:
            st.success("✅ AI 辨識完成！請確認或修改下方基本資料與匯款帳戶欄位：")
            
            with st.form("vendor_edit_form"):
                current_data = st.session_state[file_key]
                
                st.markdown("##### 🏢 基本資料")
                c1, c2 = st.columns(2)
                with c1:
                    vendor_name = st.text_input("公司/供應商名稱", value=str(current_data.get("vendor_name") or ""))
                    contact_person = st.text_input("聯絡人", value=str(current_data.get("contact_person") or ""))
                    email = st.text_input("電子郵件", value=str(current_data.get("email") or ""))
                with c2:
                    tax_id = st.text_input("統一編號", value=str(current_data.get("tax_id") or ""))
                    phone = st.text_input("聯絡電話", value=str(current_data.get("phone") or ""))
                    address = st.text_input("地址", value=str(current_data.get("address") or ""))
                
                service_scope = st.text_input("🎯 服務範疇 / 營業項目", value=str(current_data.get("service_scope") or ""), help="例如：IT軟體開發、廣告行銷、辦公用品供應等")
                
                st.markdown("---")
                st.markdown("##### 🏦 匯款帳戶資訊")
                b1, b2 = st.columns(2)
                with b1:
                    bank_name = st.text_input("解款銀行名稱", value=str(current_data.get("bank_name") or ""))
                    bank_code = st.text_input("銀行/分行代碼", value=str(current_data.get("bank_code") or ""))
                    bank_account_number = st.text_input("匯款帳號", value=str(current_data.get("bank_account_number") or ""))
                with b2:
                    branch_name = st.text_input("分行名稱", value=str(current_data.get("branch_name") or ""))
                    bank_account_name = st.text_input("戶名", value=str(current_data.get("bank_account_name") or ""))
                
                st.markdown("<br>", unsafe_allow_html=True)
                submit_button = st.form_submit_button("💾 確認資料無誤，寫入雲端『供應商資料主檔.xlsx』")
                
                if submit_button:
                    final_data = {
                        "vendor_name": vendor_name,
                        "tax_id": tax_id,
                        "service_scope": service_scope,
                        "contact_person": contact_person,
                        "phone": phone,
                        "email": email,
                        "address": address,
                        "bank_name": bank_name,
                        "branch_name": branch_name,
                        "bank_code": bank_code,
                        "bank_account_name": bank_account_name,
                        "bank_account_number": bank_account_number
                    }
                    with st.spinner("正寫入雲端資料庫..."):
                        success, msg = update_vendor_excel_on_drive(final_data, cleaned_folder_id)
                        if success:
                            st.balloons()
                            st.success("🎉 資料已成功寫入 Google Drive 雲端主檔 Excel！")
                        else:
                            st.error(f"❌ {msg}")

st.markdown("---")
st.subheader("📊 雲端『供應商資料主檔.xlsx』線上檢視")

if cleaned_folder_id:
    try:
        drive_service = get_drive_service()
        if drive_service:
            q = f"'{cleaned_folder_id}' in parents and name='供應商資料主檔.xlsx' and trashed=false"
            res = drive_service.files().list(
                q=q,
                fields="files(id, name)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            files = res.get("files", [])
            if files:
                f_id = files[0]["id"]
                req = drive_service.files().get_media(fileId=f_id)
                fh = BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                df_main = pd.read_excel(fh)
                st.dataframe(df_main, use_container_width=True)
            else:
                st.info("目前雲端尚未建立『供應商資料主檔.xlsx』，上傳寫入資料後將自動生成。")
    except Exception as ex:
        st.warning(f"線上檢視載入失敗: {ex}")
