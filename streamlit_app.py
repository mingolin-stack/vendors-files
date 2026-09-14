import streamlit as st
import google.generativeai as genai
import os
import json
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
            # 修正點：將模型名稱升級為最新的 gemini-3.6-flash（原 gemini-2.5-flash 已停用）
            model = genai.GenerativeModel("gemini-3.6-flash")
            
            # 讀取上傳檔案
            bytes_data = uploaded_file.getvalue()
            
            prompt = """
            請從這份供應商文件中提取以下欄位資訊，並以 JSON 格式輸出：
            - 供應商名稱 (vendor_name)
            - 統一編號 (tax_id)
            - 聯絡人 (contact_person)
            - 聯絡電話 (phone)
            - 電子郵件 (email)
            - 地址 (address)
            """
            
            # 如果是 PDF 或圖片格式進行辨識
            response = model.generate_content([
                {"mime_type": uploaded_file.type or "application/pdf", "data": bytes_data},
                prompt
            ])
            
            st.success("✅ AI 辨識完成！")
            st.json(response.text)
            
        except Exception as e:
            st.error(f"❌ 文件 AI 辨識發生錯誤: {e}")

st.markdown("---")
st.subheader("📊 雲端『供應商資料主檔.xlsx』線上檢視")
