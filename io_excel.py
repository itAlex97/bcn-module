import os
import pandas as pd

SHEET_CHARTED = "Charted_BOM_CapXC"
SHEET_MFGPRO = "MFGPro_CAPH"
SHEET_COMPONENTS = "Data"


def _resolve_excel_path(path):
    if os.path.exists(path):
        return path
    if not (path.lower().endswith(".xlsx") or path.lower().endswith(".xls")):
        for ext in (".xlsx", ".xls"):
            candidate = f"{path}{ext}"
            if os.path.exists(candidate):
                return candidate
    raise FileNotFoundError(f"Archivo no encontrado: {path}")


def load_excel_file(path):
    path = _resolve_excel_path(path)
    excel_file = pd.ExcelFile(path)
    nombres = excel_file.sheet_names
    if SHEET_CHARTED not in nombres:
        raise ValueError(f"No se encontro la pestaña '{SHEET_CHARTED}'.")
    if SHEET_MFGPRO not in nombres:
        raise ValueError(f"No se encontro la pestaña '{SHEET_MFGPRO}'.")
    return excel_file


def read_charted_raw(excel_file):
    return pd.read_excel(excel_file, sheet_name=SHEET_CHARTED, skiprows=8, header=0)


def read_mfgpro_raw(excel_file):
    return pd.read_excel(excel_file, sheet_name=SHEET_MFGPRO, skiprows=3, header=0)


def load_components_df(path, sheet_name=SHEET_COMPONENTS):
    path = _resolve_excel_path(path)
    return pd.read_excel(path, sheet_name=sheet_name, header=0)
