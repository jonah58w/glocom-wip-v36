import io
import os
import shutil
import subprocess
import tempfile
import pandas as pd
from utils import normalize_columns


def try_read_excel_bytes(file_bytes: bytes, header=0, sheet_name=0):
    bio = io.BytesIO(file_bytes)
    return pd.read_excel(bio, header=header, sheet_name=sheet_name)


def convert_xls_bytes_to_xlsx_bytes(file_bytes: bytes, original_name: str) -> bytes:
    suffix = os.path.splitext(original_name)[1].lower()
    if suffix != ".xls":
        return file_bytes
    with tempfile.TemporaryDirectory() as td:
        src_path = os.path.join(td, original_name)
        with open(src_path, "wb") as f:
            f.write(file_bytes)
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise RuntimeError("無法轉換 .xls：系統未安裝 libreoffice / soffice，且 pandas 讀取 .xls 也失敗。")
        cmd = [soffice, "--headless", "--convert-to", "xlsx", "--outdir", td, src_path]
        result = subprocess.run(cmd, capture_output=True, text=True)
        xlsx_path = os.path.join(td, os.path.splitext(original_name)[0] + ".xlsx")
        if result.returncode != 0 or not os.path.exists(xlsx_path):
            raise RuntimeError(f".xls 轉檔失敗：{result.stderr or result.stdout or 'unknown error'}")
        with open(xlsx_path, "rb") as f:
            return f.read()


def get_excel_file_obj(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    name = uploaded_file.name
    try:
        return pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception:
        if name.lower().endswith(".xls"):
            xlsx_bytes = convert_xls_bytes_to_xlsx_bytes(file_bytes, name)
            return pd.ExcelFile(io.BytesIO(xlsx_bytes))
        raise


def read_first_nonempty_sheet_raw(uploaded_file):
    xls = get_excel_file_obj(uploaded_file)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if not df.empty and df.dropna(how="all").shape[0] > 0:
            return df, sheet
    return pd.DataFrame(), None


def read_first_nonempty_sheet_with_header(uploaded_file, header=0):
    xls = get_excel_file_obj(uploaded_file)
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header)
            df = normalize_columns(df)
            if not df.empty and df.dropna(how="all").shape[0] > 0 and df.shape[1] > 0:
                return df, sheet
        except Exception:
            continue
    return pd.DataFrame(), None
