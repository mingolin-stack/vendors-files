import io
import pandas as pd
from typing import Dict, Any, List

def json_to_excel_bytes(extracted_data: Dict[str, Any], sheet_name: str = "供應商資料主檔") -> bytes:
    """
    將 AI 辨識回傳的 JSON / Dict 資料轉換為 Excel 檔案 (bytes)
    
    :param extracted_data: AI 辨識出的 JSON 資料字典
    :param sheet_name: Excel 工作表名稱
    :return: 包含 Excel 內容的 bytes
    """
    # 展平字典結構（若有巢狀 JSON 則拉平）
    flat_data = {}
    for key, value in extracted_data.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                flat_data[f"{key}_{sub_key}"] = sub_value
        elif isinstance(value, list):
            flat_data[key] = ", ".join(map(str, value))
        else:
            flat_data[key] = value

    # 轉為 Pandas DataFrame (單筆記錄列)
    df = pd.DataFrame([flat_data])

    # 輸出至記憶體 bytes 流
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    
    return output.getvalue()


def append_json_to_existing_excel_bytes(
    existing_excel_bytes: bytes, 
    new_data: Dict[str, Any], 
    sheet_name: str = "供應商資料主檔"
) -> bytes:
    """
    將新的辨識資料追加 (Append) 到既有的 Excel 檔案中
    
    :param existing_excel_bytes: 既有 Excel 的 bytes
    :param new_data: 新的辨識 JSON 資料
    :param sheet_name: Excel 工作表名稱
    :return: 更新後的 Excel bytes
    """
    # 展平新資料
    flat_data = {}
    for key, value in new_data.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                flat_data[f"{key}_{sub_key}"] = sub_value
        elif isinstance(value, list):
            flat_data[key] = ", ".join(map(str, value))
        else:
            flat_data[key] = value

    # 讀取既有 Excel 內容
    try:
        existing_df = pd.read_excel(io.BytesIO(existing_excel_bytes), sheet_name=sheet_name)
    except Exception:
        existing_df = pd.DataFrame()

    new_df = pd.DataFrame([flat_data])

    # 合併新舊資料
    updated_df = pd.concat([existing_df, new_df], ignore_index=False, sort=False)

    # 輸出更新後的 Excel bytes
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        updated_df.to_excel(writer, index=False, sheet_name=sheet_name)

    return output.getvalue()


def excel_bytes_to_json(excel_bytes: bytes, sheet_name: str = "供應商資料主檔") -> List[Dict[str, Any]]:
    """
    讀取 Excel 檔案內容並轉換回 JSON List 格式
    
    :param excel_bytes: Excel 的 bytes
    :param sheet_name: 工作表名稱
    :return: JSON 陣列資料
    """
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name)
    # 將 NaN 轉為 None/null 以利 JSON 處理
    df = df.where(pd.notnull(df), None)
    return df.to_dict(orient='records')
