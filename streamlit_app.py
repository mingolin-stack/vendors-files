import os
import io
import json
import base64
import fitz  # PyMuPDF
import docx
import pandas as pd
import PIL.Image
import google.generativeai as genai
import streamlit as st
from drive_utils import (
    get_drive_service,
    upload_file_from_bytes,
    search_file,
    download_file_bytes,
    upload_or_update_xlsx
)

# --------------------------------------------------
# 1. 主檔處理邏輯
# --------------------------------------------------
def process_master_excel(old_excel_bytes: bytes, new_data: dict) -> bytes:
    columns = [
        "公司全名", "統一編號", "負責人", "聯絡人",
        "公司電話", "聯絡電話", "聯絡地址", "帳單地址",
        "匯款帳號戶名", "匯款銀行", "分行別", "匯款帳號"
    ]

    if old_excel_bytes:
        try:
            df = pd.read_excel(io.BytesIO(old_excel_bytes))
        except Exception:
            df = pd.DataFrame(columns=columns)
    else:
        df = pd.DataFrame(columns=columns)

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    new_row = {col: new_data.get(col, "") for col in columns}
    new_df = pd.DataFrame([new_row])

    tax_id = str(new_data.get("統一編號", "")).strip() if new_data.get("統一編號") else ""
    
    if tax_id and "統一編號" in df.columns and (df["統一編號"].astype(str).str.strip() == tax_id).any():
        idx = df[df["統一編號"].astype(str).str.strip() == tax_id].index[0]
        for col in columns:
            if new_data.get(col):
                df.at[idx, col] = new_data[col]
    else:
        df = pd.concat([df, new_df], ignore_index=True)

    df = df[columns]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="供應商主檔")
    
    return output.getvalue()


# --------------------------------------------------
# 2. 文件內容提取與 Gemini AI 辨識邏輯
# --------------------------------------------------
def extract_content_from_file(file_bytes: bytes, file_ext: str):
    file_ext = file_ext.lower()
    
    if file_ext == ".pdf":
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        if len(doc) == 0:
            raise ValueError("上傳的 PDF 檔案為空檔")
        page = doc[0]
        pix = page.get_pixmap(dpi=200)
        return "image", pix.tobytes("png")

    elif file_ext in [".docx", ".doc"]:
        doc = docx.Document(io.BytesIO(file_bytes))
        full_text = []
        for p in doc.paragraphs:
            if p.text.strip():
                full_text.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_data:
                    full_text.append(" | ".join(row_data))
        return "text", "\n".join(full_text)

    elif file_ext in [".xlsx", ".xls"]:
        df_dict = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
        all_sheets_text = []
        for sheet_name, df in df_dict.items():
            all_sheets_text.append(f"--- Sheet: {sheet_name} ---")
            all_sheets_text.append(df.to_csv(index=False))
        return "text", "\n".join(all_sheets_text)

    else:
        raise ValueError(f"不支援的檔案格式: {file_ext}")


def analyze_vendor_document(file_bytes: bytes, file_ext: str) -> dict:
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("未設定 GEMINI_API_KEY，請在 Streamlit Secrets 設定中新增。")

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    content_type, extracted_content = extract_content_from_file(file_bytes, file_ext)

    prompt = """
    請幫我辨識這份供應商資料文件的內容，並以 JSON 格式回傳以下欄位：
    - 公司全名
    - 統一編號
    - 負責人
    - 聯絡人
    - 公司電話
    - 聯絡電話
    - 聯絡地址
    - 帳單地址
    - 匯款帳號戶名
    - 匯款銀行
    - 分行別
    - 匯款帳號

    若欄位不存在或無法辨識，請填寫 ""。僅回傳純 JSON 字串，不要包含任何 markdown 標籤。
    """

    if content_type == "image":
        image = PIL.Image.open(io.BytesIO(extracted_content))
        response = model.generate_content([prompt, image])
    else:
        user_message = f"{prompt}\n\n以下是文件的內容資料：\n{extracted_content}"
        response = model.generate_content(user_message)

    result_text = response.text.strip()
    if result_text.startswith("```json"):
        result_text = result_text[7:]
    if result_text.endswith("```"):
        result_text = result_text[:-3]

    return json.loads(result_text.strip())


# --------------------------------------------------
# 3. Streamlit 主頁面介面
# --------------------------------------------------
st.set_page_config(page_title="供應商資料辨識與主檔管理", layout="wide")

st.title("📋 供應商資料自動辨識與主檔系統 (Gemini AI 版)")
st.write("支援上傳格式：**PDF (.pdf)**, **Word (.docx / .doc)**, **Excel (.xlsx / .xls)**")

try:
    service = get_drive_service()
    drive_folder_id = st.secrets.get("drive_folder_id")
except Exception as e:
    st.error(f"Google Drive API 連線失敗，請檢查 Secrets 設定: {e}")
    st.stop()

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
            with st.spinner("Gemini AI 正在解析文件內容與提取欄位..."):
                extracted_data = analyze_vendor_document(file_bytes, file_ext)

            st.success("✅ AI 欄位解析完成！")
            st.json(extracted_data)

            if st.button("💾 將解析結果寫入雲端『供應商資料主檔.xlsx』"):
                with st.spinner("正在更新主檔並上傳至 Google Drive..."):
                    master_file_name = "供應商資料主檔.xlsx"
                    master_file_info = search_file(service, drive_folder_id, master_file_name)

                    old_excel_bytes = None
                    if master_file_info:
                        old_excel_bytes = download_file_bytes(service, master_file_info["id"])

                    updated_excel_bytes = process_master_excel(old_excel_bytes, extracted_data)
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
