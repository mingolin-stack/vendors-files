import base64
import json
import streamlit as st
from openai import OpenAI

def encode_image_to_base64(image_bytes: bytes) -> str:
    """將圖片 bytes 轉換為 Base64 字串"""
    return base64.b64encode(image_bytes).decode('utf-8')

def extract_data_from_image(image_bytes: bytes, mime_type: str, master_schema: dict) -> dict:
    """
    使用 Vision LLM (GPT-4o) 讀取表單圖片，並根據 master_schema 的欄位定義自動擷取寫入資料。

    :param image_bytes: 上傳圖片的位元組資料
    :param mime_type: 圖片 MIME 類型 (例如 'image/png', 'image/jpeg')
    :param master_schema: 定義好的欄位 JSON 結構範本
    :return: 擷取對應後的 JSON 資料
    """
    # 從 secrets 取得 API 金鑰
    api_key = st.secrets.get("OPENAI_API_KEY")
    if not api_key:
        st.error("❌ 未設定 OPENAI_API_KEY，請在 Streamlit Secrets 設定中新增。")
        return {}

    client = OpenAI(api_key=api_key)
    base64_image = encode_image_to_base64(image_bytes)

    prompt = f"""
    你是一個專業的文件與表單欄位辨識專家。
    請詳細分析這張表單圖片，讀取使用者在空白處填寫的內容。

    請務必回傳嚴格符合以下結構的 JSON 格式，不要包含 Markdown 標記或額外說明：
    {json.dumps(master_schema, ensure_ascii=False, indent=2)}

    說明規範：
    1. 尋找圖片中與 Schema 欄位標籤名稱對應的填寫文字。
    2. 若欄位空白、無法辨識或未填寫，請將其值填入 null。
    3. 日期欄位請格式化為 YYYY-MM-DD（若可解析）。
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{base64_image}"
                            },
                        },
                    ],
                }
            ],
            response_format={"type": "json_object"},
            temperature=0.1
        )

        result_text = response.choices[0].message.content
        return json.loads(result_text)

    except Exception as e:
        st.error(f"❌ 圖片 AI 辨識發生錯誤: {e}")
        return {}
