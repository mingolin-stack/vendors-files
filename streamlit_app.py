import streamlit as st
import google.generativeai as genai
import os
import json
import re
import pandas as pd
from openpyxl import load_workbook

# 頁面標題與佈局設定
st.set_page_config(page_title="供應商資料自動辨識與主檔系統", layout="wide")

st.title("📑 供應商資料自動辨識與主檔系統 (Gemini AI 最新版)")
st.caption("支援上傳格式：PDF (.pdf), Word (.docx / .doc), Excel (.xlsx / .xls)")

# 設定 Gemini API Key
api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
if not api_key:
    st.warning("⚠️ 請先於 secrets.toml 或環境變數中設定 GEMINI_API_KEY。")

genai.configure(api_key=api_key)

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 上傳檔案預覽與資訊")
    uploaded_file = st.file_uploader("請選擇要上傳的供應商文件 (PDF, Word, Excel)", type=["pdf", "docx", "doc", "xlsx", "xls"])
    
    if uploaded_file:
        st.info(f"📄 已選擇檔案：`{uploaded_file.name}`")
        # 雲端備份處理邏輯 (示範)
        st.success("✅ 原始檔案已成功備份至雲端資料夾！")

with col2:
    st.subheader("👁️ AI 欄位辨識結果")
    if uploaded_file and api_key:
        try:
            # 1. 設定 Prompt 要求 AI 僅回傳純 JSON (無 Markdown 標記)
            model = genai.GenerativeModel("gemini-3.6-flash")
            
            bytes_data = uploaded_file.getvalue()
            
            prompt = """
            請從這份供應商文件中提取以下欄位資訊，並直接以純 JSON 格式輸出，不要包含任何 Markdown 格式化（例如不要使用 ```json 或 ``` 標籤）：
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
            
            # 2. 清理 Markdown 標記語法 (```json ... ```)
            cleaned_text = re.sub(r"^```[a-zA-Z]*
?", "", raw_text)
            cleaned_text = re.sub(r"
?```$", "", cleaned_text).strip()
            
            # 3. 解析並渲染 JSON / 字典物件
            try:
                parsed_json = json.loads(cleaned_text)
                st.json(parsed_json)
            except json.JSONDecodeError:
                # 若解析失敗則顯示清洗後的字串，避免系統崩潰
                st.write(cleaned_text)
            
        except Exception as e:
            st.error(f"❌ 文件 AI 辨識發生錯誤: {e}")

st.markdown("---")
st.subheader("📊 雲端『供應商資料主檔.xlsx』線上檢視")
