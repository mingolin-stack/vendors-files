import os
import io
import mimetypes
from typing import Optional, Dict, Any, List
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload, MediaIoBaseDownload

NUM_RETRIES = 3
SCOPES = ['https://www.googleapis.com/auth/drive']


def handle_http_error(error: HttpError, action_description: str) -> None:
    """解析並顯示 Google Drive API 錯誤"""
    status_code = error.resp.status
    try:
        error_details = error.content.decode("utf-8")
    except Exception:
        error_details = str(error)

    st.error(f"❌ {action_description} 失敗 [HTTP Status: {status_code}]")
    st.code(error_details, language="json")


def get_drive_service():
    """從 st.secrets 初始化並取得 Google Drive API Service 物件"""
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
        else:
            creds_dict = dict(st.secrets)

        credentials = Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
        service = build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        st.error(f"初始化 Google Drive Service 失敗: {e}")
        raise e


def find_or_create_folder(service, parent_id: str, folder_name: str) -> Optional[str]:
    """尋找指定父資料夾下的同名資料夾；若不存在則建立"""
    safe_folder_name = folder_name.replace("'", "\\'")
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{safe_folder_name}' and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )

    try:
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


def find_file_id(service, parent_id: str, file_name: str) -> Optional[str]:
    """在指定資料夾中搜尋檔案 ID"""
    safe_file_name = file_name.replace("'", "\\'")
    query = (
        f"'{parent_id}' in parents and "
        f"name = '{safe_file_name}' and "
        f"trashed = false"
    )
    try:
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
        return files[0]["id"] if files else None
    except HttpError as error:
        handle_http_error(error, f"搜尋檔案 '{file_name}'")
        raise error


def download_file_bytes(service, file_id: str) -> bytes:
    """根據 file_id 下載雲端檔案並回傳 bytes 資料"""
    try:
        request = service.files().get_media(fileId=file_id)
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return file_stream.getvalue()
    except HttpError as error:
        handle_http_error(error, f"下載檔案 (ID: {file_id})")
        raise error


def upload_or_update_xlsx(service, parent_folder_id: str, file_name: str, file_bytes: bytes) -> str:
    """若指定名稱的 XLSX 存在則更新，若不存在則新建檔案"""
    existing_file_id = find_file_id(service, parent_folder_id, file_name)
    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )

    try:
        if existing_file_id:
            # 存在舊檔 -> 更新內容
            updated_file = (
                service.files()
                .update(
                    fileId=existing_file_id,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True
                )
                .execute(num_retries=NUM_RETRIES)
            )
            return updated_file.get("id")
        else:
            # 不存在舊檔 -> 新增檔案
            file_metadata = {
                "name": file_name,
                "parents": [parent_folder_id]
            }
            new_file = (
                service.files()
                .create(
                    body=file_metadata,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True
                )
                .execute(num_retries=NUM_RETRIES)
            )
            return new_file.get("id")
    except HttpError as error:
        handle_http_error(error, f"上傳/更新 Excel 檔案 '{file_name}'")
        raise error