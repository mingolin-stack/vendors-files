import streamlit as st
import google.generativeai as genai
import os
import json
import re
import pandas as pd
import time
import threading
import uvicorn
from io import BytesIO
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ---------------------------------------------------------
# 1. 初始化與 API Key 設定
# ---------------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
drive_folder_id = st.secrets.get("drive_folder_id") or os.environ.get("DRIVE_FOLDER_ID")

if api_key:
    genai.configure(api_key=api_key)

PREFERRED_MODELS = [
    "gemini-1.5-flash",
    "gemini-1.5-flash-latest",
    "gemini-1.5-pro",
    "gemini-2.5-flash",
    "gemini-3.6-flash"
]

def get_candidate_models():
    try:
        available_models = [
            m.name.replace("models/", "") for m in genai.list_models() 
            if 'generateContent' in m.supported_generation_methods and 'deprecated' not in m.name.lower()
        ]
        candidates = [p for p in PREFERRED_MODELS if p in available_models]
        for m in available_models:
            if m not in candidates and 'flash' in m.lower():
                candidates.append(m)
        for m in available_models:
            if m not in candidates:
                candidates.append(m)
        return candidates
    except Exception:
        return PREFERRED_MODELS

def extract_vendor_info_from_bytes(bytes_data: bytes, mime_type: str) -> dict:
    """呼叫 Gemini AI 進行文件辨識並傳回 JSON 字典"""
    candidate_models = get_candidate_models()
    prompt = """
    請從這份供應商文件中提取以下欄位資訊，並直接以純 JSON 格式輸出：
    {
      "vendor_name": "供應商名稱",
      "tax_id": "統一編號",
      "service_scope": "服務範疇",
      "contact_person": "聯絡人",
      "phone": "聯絡電話",
      "email": "電子郵件",
      "address": "地址",
      "bank_name": "銀行名稱",
      "branch_name": "分行名稱",
      "bank_code": "銀行代碼",
      "bank_account_name": "戶名",
      "bank_account_number": "匯款帳號"
    }
    如欄位不存在請留空字串 ""。
    """
    last_error = None
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content([
                {"mime_type": mime_type, "data": bytes_data},
                prompt
            ])
            raw_text = response.text.strip()
            cleaned_text = re.sub(r"^```[a-zA-Z]*", "", raw_text)
            cleaned_text = re.sub(r"```$", "", cleaned_text).strip()
            return json.loads(cleaned_text)
        except Exception as e:
            last_error = e
            time.sleep(1)
            continue
            
    raise last_error if last_error else Exception("AI 辨識失敗")

# ---------------------------------------------------------
# 2. 開啟 FastAPI (讓 ChatGPT 呼叫 API)
# ---------------------------------------------------------
api_app = FastAPI(title="Vendor Extraction API", description="供 ChatGPT Actions 呼叫的文件辨識 API")

api_app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@api_app.post("/api/extract_vendor")
async def api_extract_vendor(file: UploadFile = File(...)):
    """API 端點：接受 POST multipart/form-data 上傳檔案，傳回 JSON 辨識結果"""
    try:
        content = await file.read()
        mime_type = file.content_type or "application/pdf"
        result = extract_vendor_info_from_bytes(content, mime_type)
        return {"status": "success", "data": result}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def run_fastapi():
    uvicorn.run(api_app, host="0.0.0.0", port=8000)

# 在背景執行 FastAPI
if "fastapi_thread" not in st.session_state:
    thread = threading.Thread(target=run_fastapi, daemon=True)
    thread.start()
    st.session_state["fastapi_thread"] = True

# ---------------------------------------------------------
# 3. Streamlit 前端介面
# ---------------------------------------------------------
st.set_page_config(page_title="供應商資料自動辨識與主檔系統", layout="wide")

st.title("📑 供應商資料自動辨識與主檔系統 (Gemini AI + ChatGPT API 支援)")
st.caption("支援網頁上傳與 ChatGPT API 遠端呼叫 (`POST http://<YOUR_HOST>:8000/api/extract_vendor`)")

# 前端網頁介面邏輯...
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 上傳檔案預覽與資訊")
    uploaded_file = st.file_uploader("選擇供應商文件", type=["pdf", "docx", "doc", "xlsx", "xls"])
    if uploaded_file:
        st.info(f"📄 已選擇檔案：`{uploaded_file.name}`")

with col2:
    st.subheader("👁️ AI 欄位辨識與核對編輯")
    if uploaded_file and api_key:
        file_key = f"parsed_scope_{uploaded_file.name}"
        if file_key not in st.session_state:
            with st.spinner("AI 正在辨識文件..."):
                try:
                    bytes_data = uploaded_file.getvalue()
                    mime_type = uploaded_file.type or "application/pdf"
                    parsed_json = extract_vendor_info_from_bytes(bytes_data, mime_type)
                    st.session_state[file_key] = parsed_json
                except Exception as e:
                    st.error(f"❌ 文件 AI 辨識發生錯誤: {e}")
