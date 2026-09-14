"""
供應商資料表 PDF 掃描辨識工具 (Streamlit)
------------------------------------------------
流程：
  1. 上傳一份已掃描的供應商資料表 PDF
  2. 程式自動辨識(文字欄位用 Google Vision API，勾選框用像素判斷)
  3. 畫面顯示「裁切小圖 + 辨識建議值」，同仁快速核對、修正
  4. 按下確認，資料寫入：
       - 該供應商自己的主檔(Google Drive「供應商主檔」資料夾內，以統編/公司名建檔，每次提交新增一列)
       - 彙總總表(同資料夾內 彙總總表.xlsx，每家供應商固定一列，重複提交會覆蓋更新)

需要的 Streamlit secrets(在 App 設定的 Secrets 分頁貼入)：

    drive_folder_id = "你的 Google Drive 資料夾 ID"

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "mingo-lin-greenharvest-com-tw@adept-eon-496404-t4.iam.gserviceaccount.comm"
    client_id = "..."
    token_uri = "https://oauth2.googleapis.com/token"

    (把下載到的 .json 金鑰檔內容，轉成上面這種 TOML 格式貼進去即可；
     private_key 裡的換行請保留 \n)
------------------------------------------------
"""

import json
from io import BytesIO

import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
from google.cloud import vision

from drive_utils import get_drive_service, find_or_create_folder, find_file_id, download_file_bytes, upload_or_update_xlsx
from vision_utils import crop_field, ocr_text, is_checked
from master_utils import build_columns, record_filename, append_row_to_workbook, upsert_row_in_summary, build_row_dict

st.set_page_config(page_title="供應商資料表 PDF 辨識工具", page_icon="🧾", layout="wide")

RENDER_DPI = 200  # 必須跟 template.json 校正時使用的 DPI 一致


@st.cache_data
def load_template():
    with open("template.json", encoding="utf-8") as f:
        return json.load(f)


@st.cache_resource
def get_vision_client():
    creds_info = dict(st.secrets["gcp_service_account"])
    from google.oauth2 import service_account
    creds = service_account.Credentials.from_service_account_info(creds_info)
    return vision.ImageAnnotatorClient(credentials=creds)


@st.cache_resource
def get_drive():
    creds_info = dict(st.secrets["gcp_service_account"])
    return get_drive_service(creds_info)


def pdf_to_image(pdf_bytes: bytes, dpi: int = RENDER_DPI) -> Image.Image:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]  # 假設表格都在第一頁
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


def run_extraction(page_image, template, vision_client):
    """對整份表格跑一次辨識，回傳 {欄位名: 建議值} 的字典，供畫面顯示與人工核對。"""
    suggestions = {}
    crops = {}
    for field in template["fields"]:
        if field["type"] == "text":
            crop = crop_field(page_image, field["box"])
            crops[field["name"]] = crop
            try:
                suggestions[field["name"]] = ocr_text(vision_client, crop)
            except Exception as e:
                suggestions[field["name"]] = ""
                st.warning(f"「{field['name']}」辨識失敗：{e}")
        elif field["type"] == "checkbox_single":
            crop = crop_field(page_image, field["box"])
            crops[field["name"]] = crop
            checked, ratio = is_checked(page_image, field["box"])
            suggestions[field["name"]] = "是" if checked else "否"
        elif field["type"] == "checkbox_group":
            best_label, best_ratio = None, 0
            option_crops = []
            for opt in field["options"]:
                crop = crop_field(page_image, opt["box"])
                option_crops.append((opt["label"], crop))
                checked, ratio = is_checked(page_image, opt["box"])
                if checked and ratio > best_ratio:
                    best_label, best_ratio = opt["label"], ratio
            crops[field["name"]] = option_crops
            suggestions[field["name"]] = best_label or ""
    return suggestions, crops


def main():
    st.title("🧾 供應商資料表 PDF 掃描辨識工具")
    st.caption("上傳掃描好的供應商資料表 PDF，自動辨識後請核對，確認無誤再存檔。")

    with st.expander("🔧 Secrets 診斷工具(排除問題用，確認沒問題後可以刪掉這段)"):
        try:
            info = dict(st.secrets["gcp_service_account"])
            pk = info.get("private_key", "")
            st.write("client_email:", repr(info.get("client_email", "")))
            st.write("project_id:", repr(info.get("project_id", "")))
            st.write("private_key_id:", repr(info.get("private_key_id", "")))
            st.write("client_id:", repr(info.get("client_id", "")))
            st.write("token_uri:", repr(info.get("token_uri", "")))
            st.write("private_key 開頭 20 字元:", repr(pk[:20]))
            st.write("private_key 結尾 20 字元:", repr(pk[-20:]))
            st.write("private_key 總長度:", len(pk))
            st.write("private_key 裡「真的換行符號」數量:", pk.count("\n"))
            st.write("private_key 裡「反斜線+n 兩個字元」數量:", pk.count("\\n"))
            st.write("drive_folder_id:", repr(st.secrets.get("drive_folder_id", "")))
        except Exception as e:
            st.error(f"讀取 Secrets 時發生錯誤：{e}")

    template = load_template()
    columns = build_columns(template)

    uploaded_pdf = st.file_uploader("上傳已掃描的供應商資料表(PDF，一次一份)", type=["pdf"])

    if uploaded_pdf is not None:
        pdf_bytes = uploaded_pdf.getvalue()

        if st.session_state.get("_current_file") != uploaded_pdf.name:
            # 換了一份新檔案，清掉之前的暫存結果
            st.session_state["_current_file"] = uploaded_pdf.name
            st.session_state.pop("_suggestions", None)
            st.session_state.pop("_crops", None)

        if "_suggestions" not in st.session_state:
            with st.spinner("辨識中，請稍候(掃描頁面較多或網路較慢時可能需要一點時間)..."):
                page_image = pdf_to_image(pdf_bytes)
                vision_client = get_vision_client()
                suggestions, crops = run_extraction(page_image, template, vision_client)
                st.session_state["_suggestions"] = suggestions
                st.session_state["_crops"] = crops
                st.session_state["_page_image"] = page_image

        st.success("辨識完成，請逐一核對下方內容，確認或修正後再送出。")
        st.divider()

        edited = {}
        for field in template["fields"]:
            name = field["name"]
            col_img, col_val = st.columns([1, 2])

            if field["type"] == "text":
                with col_img:
                    st.image(st.session_state["_crops"][name], use_container_width=True)
                with col_val:
                    edited[name] = st.text_input(name, value=st.session_state["_suggestions"][name], key=f"in_{name}")

            elif field["type"] == "checkbox_single":
                with col_img:
                    st.image(st.session_state["_crops"][name], use_container_width=True)
                with col_val:
                    default_yes = st.session_state["_suggestions"][name] == "是"
                    edited[name] = "是" if st.checkbox(name, value=default_yes, key=f"in_{name}") else "否"

            elif field["type"] == "checkbox_group":
                option_labels = [o["label"] for o in field["options"]]
                suggested = st.session_state["_suggestions"][name]
                with col_img:
                    thumb_cols = st.columns(len(st.session_state["_crops"][name]))
                    for tc, (label, crop) in zip(thumb_cols, st.session_state["_crops"][name]):
                        with tc:
                            st.image(crop, caption=label, use_container_width=True)
                with col_val:
                    default_idx = option_labels.index(suggested) if suggested in option_labels else 0
                    edited[name] = st.radio(name, option_labels, index=default_idx, key=f"in_{name}", horizontal=True)

            st.divider()

        if st.button("✅ 確認並存檔", type="primary"):
            with st.spinner("寫入 Google Drive 中..."):
                drive_folder_id = st.secrets["drive_folder_id"]
                service = get_drive()

                supplier_folder_id = find_or_create_folder(service, drive_folder_id, "供應商主檔")

                row = build_row_dict(uploaded_pdf.name, edited)

                # 1) 該供應商自己的主檔(新增一列，形成歷史紀錄)
                fname = record_filename(edited)
                existing_id = find_file_id(service, supplier_folder_id, fname)
                existing_bytes = download_file_bytes(service, existing_id) if existing_id else None
                new_bytes = append_row_to_workbook(existing_bytes, columns, row)
                upload_or_update_xlsx(service, supplier_folder_id, fname, new_bytes)

                # 2) 彙總總表(同一家供應商覆蓋更新那一列)
                summary_name = "彙總總表.xlsx"
                summary_id = find_file_id(service, drive_folder_id, summary_name)
                summary_bytes = download_file_bytes(service, summary_id) if summary_id else None
                new_summary_bytes = upsert_row_in_summary(summary_bytes, columns, row)
                upload_or_update_xlsx(service, drive_folder_id, summary_name, new_summary_bytes)

            st.success(f"已存檔！供應商主檔：{fname}，並已同步更新彙總總表。")
            st.balloons()
            for k in ["_suggestions", "_crops", "_page_image", "_current_file"]:
                st.session_state.pop(k, None)


if __name__ == "__main__":
    main()
