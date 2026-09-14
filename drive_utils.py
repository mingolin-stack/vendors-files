import os
import io
import mimetypes
from typing import Optional, Dict, Any, List
import streamlit as st
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload

# 預設失敗重試次數
NUM_RETRIES = 3


def handle_http_error(error: HttpError, action_description: str) -> None:
    """解析並在 Streamlit 上顯示詳細的 Google Drive API 錯誤訊息"""
    status_code = error.resp.status
    try:
        error_details = error.content.decode("utf-8")
    except Exception:
        error_details = str(error)

    st.error(f"❌ {action_description} 失敗 [HTTP Status: {status_code}]")
    st.code(error_details, language="json")

    if status_code == 404:
        st.warning("提示：請確認目標資料夾/檔案 ID 是否正確，且該項目確實存在。")
    elif status_code == 403:
        st.warning(
            "提示：權限不足！請確認已將目標資料夾「共用/共享」給 Service Account Email，並賦予「編輯者」權限。"
        )


def find_or_create_folder(service, parent_id: str, folder_name: str) -> Optional[str]:
    """尋找指定父資料夾下的同名資料夾；若不存在則自動建立。

    :param service: 已初始化的 Google Drive API service 物件
    :param parent_id: 父資料夾 ID
    :param folder_name: 要尋找或建立的資料夾名稱
    :return: 資料夾 ID
    """
    safe_folder_name = folder_name.replace("'", "\'")
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{safe_folder_name}' and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )

    try:
        # 1. 搜尋是否存在同名資料夾
        results = (
            service.files()
            .list(
                q=query,
                fields="files(id, name)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute(num_retries=NUM_RETRIES)
        )

        files = results.get("files", [])
        if files:
            return files[0]["id"]

        # 2. 若不存在，建立新資料夾
        folder_metadata = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        }
        folder = (
            service.files()
            .create(body=folder_metadata, fields="id", supportsAllDrives=True)
            .execute(num_retries=NUM_RETRIES)
        )

        return folder.get("id")

    except HttpError as error:
        handle_http_error(error, f"搜尋或建立資料夾 '{folder_name}'")
        raise error


def upload_file_from_path(
    service, local_file_path: str, parent_folder_id: str, custom_name: Optional[str] = None
) -> Optional[Dict[str, Any]]:
    """從本機檔案路徑上傳檔案至指定雲端資料夾。

    :param service: 已初始化的 Google Drive API service 物件
    :param local_file_path: 本機檔案路徑
    :param parent_folder_id: 目標雲端資料夾 ID
    :param custom_name: 上傳至雲端後的自訂檔名（若為 None 則使用原本檔名）
    :return: 包含 id, name, webViewLink 的檔案資訊字典
    """
    if not os.path.exists(local_file_path):
        st.error(f"本機檔案不存在: {local_file_path}")
        return None

    file_name = custom_name if custom_name else os.path.basename(local_file_path)
    mime_type, _ = mimetypes.guess_type(local_file_path)
    if mime_type is None:
        mime_type = "application/octet-stream"

    file_metadata = {"name": file_name, "parents": [parent_folder_id]}

    try:
        media = MediaFileUpload(local_file_path, mimetype=mime_type, resumable=True)
        uploaded_file = (
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id, name, webViewLink, webContentLink",
                supportsAllDrives=True,
            )
            .execute(num_retries=NUM_RETRIES)
        )
        return uploaded_file
    except HttpError as error:
        handle_http_error(error, f"上傳檔案 '{file_name}'")
        raise error


def upload_file_from_bytes(
    service,
    file_bytes: bytes,
    file_name: str,
    parent_folder_id: str,
    mime_type: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """從記憶體位元組 (Bytes/Streamlit UploadedFile) 上傳檔案至雲端資料夾。

    :param service: 已初始化的 Google Drive API service 物件
    :param file_bytes: 檔案內容 (bytes)
    :param file_name: 儲存在雲端的檔名
    :param parent_folder_id: 目標雲端資料夾 ID
    :param mime_type: MIME 類型（如未提供則自動推讀）
    :return: 包含 id, name, webViewLink 的檔案資訊字典
    """
    if not mime_type:
        mime_type, _ = mimetypes.guess_type(file_name)
        if mime_type is None:
            mime_type = "application/octet-stream"

    file_metadata = {"name": file_name, "parents": [parent_folder_id]}
    media_stream = io.BytesIO(file_bytes)

    try:
        media = MediaIoBaseUpload(
            media_stream, mimetype=mime_type, resumable=True
        )
        uploaded_file = (
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id, name, webViewLink, webContentLink",
                supportsAllDrives=True,
            )
            .execute(num_retries=NUM_RETRIES)
        )
        return uploaded_file
    except HttpError as error:
        handle_http_error(error, f"上傳檔案 '{file_name}'")
        raise error


def list_files_in_folder(service, parent_id: str) -> List[Dict[str, Any]]:
    """搜尋指定資料夾內的所有未刪除檔案（不包含子資料夾）。

    :param service: 已初始化的 Google Drive API service 物件
    :param parent_id: 父資料夾 ID
    :return: 檔案清單 [{"id": "...", "name": "...", "mimeType": "..."}, ...]
    """
    query = (
        f"'{parent_id}' in parents and "
        f"mimeType != 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )

    try:
        results = (
            service.files()
            .list(
                q=query,
                fields="files(id, name, mimeType, webViewLink, createdTime)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute(num_retries=NUM_RETRIES)
        )
        return results.get("files", [])
    except HttpError as error:
        handle_http_error(error, f"查詢資料夾內容 (ID: {parent_id})")
        raise error
