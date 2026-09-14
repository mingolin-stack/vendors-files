import streamlit as st
import json
import io
import mimetypes
from drive_utils import (
    get_drive_service,
    find_or_create_folder,
    upload_file_from_bytes,
    search_file,
    download_file_bytes,
    upload_or_update_xlsx,
)
from vision_utils import extract_data_from_image
from master_utils import (
    json_to_excel_bytes,
    append_json_to_existing_excel_bytes,
    excel_bytes_to_json,
)

# 頁面標題與配置
st.set_page_config(page_title="供應商資料辨識與主檔管理系統", layout="wide")

def load_master_schema():
    """載入主檔欄位結構定義范本 (template.json)"""
    default_schema = {
        "供應商名稱": None,
        "統一編號": None,
        "負責人": None,
        "聯絡電話": None,
        "電子郵件": None,
        "地址": None,
        "銀行帳號": None,
        "建立日期": None,
    }
    try:
        with open("template.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default_schema

def main():
    st.title("📄 供應商表單 / 圖片辨識與主檔自動寫入系統")
    st.markdown("支援 PDF 檔案以及 **PNG、JPG、JPEG 圖片/照片** 自動辨識並整合寫入 Google Drive Excel 主檔。")

    # 1. 初始化 Google Drive 服務與基礎參數
    try:
        service = get_drive_service()
        drive_folder_id = st.secrets.get("drive_folder_id")
        if not drive_folder_id:
            st.error("❌ 尚未在 Secrets 中設定 `drive_folder_id`。")
            st.stop()
    except Exception as e:
        st.error(f"❌ Google Drive 初始化失敗: {e}")
        st.stop()

    # 2. 準備主檔 Schema
    master_schema = load_master_schema()

    # 3. 檔案上傳區塊 (支援圖片與 PDF)
    st.subheader("1. 上傳表單或照片")
    uploaded_file = st.file_uploader(
        "請選擇或拖曳供應商文件/照片 (支援 PNG, JPG, JPEG, PDF)",
        type=["png", "jpg", "jpeg", "pdf"]
    )

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        file_name = uploaded_file.name
        mime_type = uploaded_file.type or mimetypes.guess_type(file_name)[0]

        st.divider()
        col1, col2 = st.columns([1, 1])

        # 顯示圖片預覽
        with col1:
            st.subheader("📷 上傳檔案預覽")
            if mime_type in ["image/png", "image/jpeg", "image/jpg"]:
                st.image(uploaded_file, use_container_width=True, caption=file_name)
            elif mime_type == "application/pdf":
                st.info(f"📄 已選擇 PDF 文件：{file_name}")

        # 進行雲端備份與 AI 辨識
        with col2:
            st.subheader("🤖 AI 欄位辨識結果")

            # Step 1: 自動建立/尋找資料夾並備份原檔
            with st.spinner("☁️ 正在上傳檔案備份至 Google Drive..."):
                try:
                    target_folder_id = find_or_create_folder(
                        service=service,
                        parent_folder_id=drive_folder_id,
                        folder_name="供應商主檔"
                    )
                    upload_file_from_bytes(
                        service=service,
                        file_bytes=file_bytes,
                        file_name=file_name,
                        parent_folder_id=target_folder_id,
                        mime_type=mime_type
                    )
                    st.success("✅ 原始檔案已備份至雲端資料夾！")
                except Exception as e:
                    st.error(f"❌ 檔案備份失敗: {e}")

            # Step 2: 若為圖片，進行 Vision API 自動辨識
            extracted_data = {}
            if mime_type in ["image/png", "image/jpeg", "image/jpg"]:
                with st.spinner("🔍 正在使用 AI Vision 辨識圖片內容..."):
                    extracted_data = extract_data_from_image(
                        image_bytes=file_bytes,
                        mime_type=mime_type,
                        master_schema=master_schema
                    )

                if extracted_data:
                    st.success("✅ 辨識完成！請確認下方欄位內容：")
                    # 可供使用者確認與手動修正 AI 辨識出的結果
                    edited_data = {}
                    for k, v in extracted_data.items():
                        edited_data[k] = st.text_input(label=f"【{k}】", value="" if v is None else str(v))

                    # Step 3: 確認寫入主檔 Excel 按鈕
                    if st.button("💾 確認資料並寫入/更新雲端 Excel 主檔", type="primary"):
                        with st.spinner("📊 正在更新雲端 Excel 主檔..."):
                            master_filename = "供應商資料主檔.xlsx"
                            
                            # 搜尋雲端是否已有 Excel 主檔
                            existing_file = search_file(service, target_folder_id, master_filename)

                            if existing_file:
                                file_id = existing_file["id"]
                                existing_bytes = download_file_bytes(service, file_id)
                                new_excel_bytes = append_json_to_existing_excel_bytes(
                                    existing_excel_bytes=existing_bytes,
                                    new_data=edited_data
                                )
                            else:
                                new_excel_bytes = json_to_excel_bytes(extracted_data=edited_data)

                            # 更新/上傳主檔至 Google Drive
                            upload_or_update_xlsx(
                                service=service,
                                folder_id=target_folder_id,
                                file_name=master_filename,
                                xlsx_bytes=new_excel_bytes
                            )
                            st.balloons()
                            st.success("🎉 資料已成功 append 寫入並更新至雲端『供應商資料主檔.xlsx』！")

            else:
                st.info("ℹ️ 提示：目前自動辨識功能主要針對 PNG / JPG / JPEG 圖片格式。")

    st.divider()

    # 4. 檢視雲端主檔 Excel 內容
    st.subheader("📊 雲端『供應商資料主檔.xlsx』線上檢視")
    if st.button("🔄 載入/重新整理雲端主檔資料"):
        with st.spinner("📥 正在從 Google Drive 讀取主檔..."):
            try:
                target_folder_id = find_or_create_folder(service, drive_folder_id, "供應商主檔")
                existing_file = search_file(service, target_folder_id, "供應商資料主檔.xlsx")
                if existing_file:
                    excel_bytes = download_file_bytes(service, existing_file["id"])
                    records = excel_bytes_to_json(excel_bytes)
                    st.dataframe(records, use_container_width=True)
                else:
                    st.warning("⚠️ 目前雲端上尚未存在『供應商資料主檔.xlsx』，請先上傳資料並寫入。")
            except Exception as e:
                st.error(f"❌ 讀取主檔失敗: {e}")

if __name__ == "__main__":
    main()
