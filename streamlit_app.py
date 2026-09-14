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
    # 正則提取標準 33 位元的 Google Drive Folder ID
    match = re.search(r'[a-zA-Z0-9_-]{25,50}', str(folder_id))
    if match:
        return match.group(0)
    return str(folder_id).strip().rstrip('.')

cleaned_folder_id = clean_folder_id(drive_folder_id)

def update_vendor_excel_on_drive(vendor_data, folder_id):
    service = get_drive_service()
    if not service or not folder_id:
        return False, "未設定 GCP Service Account 或 drive_folder_id"
    
    file_name = "供應商資料主檔.xlsx"
    query = f"'{folder_id}' in parents and name='{file_name}' and trashed=false"
    
    try:
        results = service.files().list(q=query, fields="files(id, name)").execute()
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
            service.files().update(fileId=file_id, media_body=media).execute()
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
            service.files().create(body=file_metadata, media_body=media, fields="id").execute()
            
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
    st.subheader("👁️ AI 欄位辨識結果")
    if uploaded_file and api_key:
        try:
            model = genai.GenerativeModel("gemini-3.6-flash")
            bytes_data = uploaded_file.getvalue()
            
            prompt = """
            請從這份供應商文件中提取以下欄位資訊，並直接以純 JSON 格式輸出：
            {
              "vendor_name": "供應商名稱",
              "tax_id": "統一編號",
              "contact_person": "聯絡人",
              "phone": "聯絡電話",
              "email": "電子郵件",
              "address": "地址"
            }
            """
            
            response = model.generate_content([
                {"mime_type": uploaded_file.type or "application/pdf", "data": bytes_data},
                prompt
            ])
            
            st.success("✅ AI 辨識完成！")
            raw_text = response.text.strip()
            
            cleaned_text = re.sub(r"^```[a-zA-Z]*\n?", "", raw_text)
            cleaned_text = re.sub(r"\n?```$", "", cleaned_text).strip()
            
            parsed_json = {}
            try:
                parsed_json = json.loads(cleaned_text)
                st.json(parsed_json)
            except json.JSONDecodeError:
                st.write(cleaned_text)
            
            if parsed_json:
                if st.button("💾 寫入雲端『供應商資料主檔.xlsx』"):
                    with st.spinner("正寫入雲端資料庫..."):
                        success, msg = update_vendor_excel_on_drive(parsed_json, cleaned_folder_id)
                        if success:
                            st.success("🎉 資料已成功寫入 Google Drive 雲端主檔 Excel！")
                        else:
                            st.error(f"❌ {msg}")

        except Exception as e:
            st.error(f"❌ 文件 AI 辨識發生錯誤: {e}")

st.markdown("---")
st.subheader("📊 雲端『供應商資料主檔.xlsx』線上檢視")

if cleaned_folder_id:
    try:
        drive_service = get_drive_service()
        if drive_service:
            q = f"'{cleaned_folder_id}' in parents and name='供應商資料主檔.xlsx' and trashed=false"
            res = drive_service.files().list(q=q, fields="files(id, name)").execute()
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
