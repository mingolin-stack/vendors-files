import os
import streamlit as st
import pandas as pd
from drive_utils import (
    get_drive_service,
    find_or_create_folder,
    upload_file_from_bytes,
    search_file,
    download_file_bytes,
    upload_or_update_xlsx
)
from vision_utils import analyze_vendor_document
from master_utils import process_master_excel

st.set_page_config(page_title="供應商資料辨識與主檔管理", layout="wide")

st.title("📋 供應商資料自動辨識與主檔系統")
st.write("支援上傳格式：**PDF (.pdf)**, **Word (.docx / .doc)**, **Excel (.xlsx / .xls)**")

# 初始化 Google Drive API 服務
try:
    service = get_drive_service()
    drive_folder_id = st.secrets.get("drive_folder_id")
except Exception as e:
    st.error(f"Google Drive API 連線失敗，請檢查 Secrets 設定: {e}")
    st.stop()

# 檔案上傳元件
uploaded_file = st.file_uploader(
    "請選擇要上傳的供應商文件 (PDF, Word, Excel)",
    type=["pdf", "docx", "doc", "xlsx", "xls"]
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_name = uploaded_file.name
    file_ext = os.path.splitext(file_name)[1].lower()

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📁 上傳檔案預覽與資訊")
        st.info(f"📄 已選擇檔案：`{file_name}`")

        # 備份原始檔案至 Google Drive
        try:
            with st.spinner("正在將原始檔案備份至 Google Drive..."):
                uploaded_drive_file = upload_file_from_bytes(
                    service=service,
                    file_bytes=file_bytes,
                    file_name=file_name,
                    parent_folder_id=drive_folder_id
                )
            st.success("✅ 原始檔案已成功備份至雲端資料夾！")
        except Exception as e:
            st.error(f"❌ 檔案上傳至 Google Drive 失敗: {e}")

    with col2:
        st.subheader("🤖 AI 欄位辨識結果")
        try:
            with st.spinner("AI 正在解析文件內容與提取欄位..."):
                extracted_data = analyze_vendor_document(file_bytes, file_ext)

            st.success("✅ AI 欄位解析完成！")
            st.json(extracted_data)

            # 更新至主檔按鈕
            if st.button("💾 將解析結果寫入雲端『供應商資料主檔.xlsx』"):
                with st.spinner("正在更新主檔並上傳至 Google Drive..."):
                    # 搜尋主檔是否存在
                    master_file_name = "供應商資料主檔.xlsx"
                    master_file_info = search_file(service, drive_folder_id, master_file_name)

                    old_excel_bytes = None
                    if master_file_info:
                        old_excel_bytes = download_file_bytes(service, master_file_info["id"])

                    # 處理並合併主檔 Excel
                    updated_excel_bytes = process_master_excel(old_excel_bytes, extracted_data)

                    # 上傳/覆蓋主檔
                    upload_or_update_xlsx(service, drive_folder_id, master_file_name, updated_excel_bytes)

                st.balloons()
                st.success("🎉 主檔更新成功！請重整底下檢視表或下載查看。")

        except Exception as e:
            st.error(f"❌ 文件 AI 辨識發生錯誤: {e}")

st.markdown("---")
st.subheader("📊 雲端『供應商資料主檔.xlsx』線上檢視")

if st.button("🔄 載入/重新整理雲端主檔資料"):
    try:
        master_file_name = "供應商資料主檔.xlsx"
        master_file_info = search_file(service, drive_folder_id, master_file_name)

        if master_file_info:
            excel_bytes = download_file_bytes(service, master_file_info["id"])
            df = pd.read_excel(pd.io.common.BytesIO(excel_bytes))
            st.dataframe(df, use_container_width=True)
            
            # 提供下載
            st.download_button(
                label="📥 下載完整『供應商資料主檔.xlsx』",
                data=excel_bytes,
                file_name="供應商資料主檔.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ 目前雲端上尚未存在『供應商資料主檔.xlsx』，請先上傳資料並寫入。")
    except Exception as e:
        st.error(f"❌ 讀取主檔失敗: {e}")
