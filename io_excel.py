import os
import pandas as pd


# Nombres fijos de hojas esperadas en los archivos Excel.
SHEET_CHARTED = "Charted_BOM_CapXC"
SHEET_MFGPRO = "MFGPro_CAPH"
SHEET_COMPONENTS = "Data"


def _resolve_excel_path(path):
    """Acepta rutas con o sin extension y confirma que el archivo exista."""
    if os.path.exists(path):
        return path
    if not (path.lower().endswith(".xlsx") or path.lower().endswith(".xls")):
        for ext in (".xlsx", ".xls"):
            candidate = f"{path}{ext}"
            if os.path.exists(candidate):
                return candidate
    raise FileNotFoundError(f"Archivo no encontrado: {path}")


def load_excel_file(path):
    """Abre el Excel principal y valida que tenga las hojas necesarias."""
    path = _resolve_excel_path(path)
    excel_file = pd.ExcelFile(path)
    nombres = excel_file.sheet_names
    if SHEET_CHARTED not in nombres:
        raise ValueError(f"No se encontro la pestana '{SHEET_CHARTED}'.")
    if SHEET_MFGPRO not in nombres:
        raise ValueError(f"No se encontro la pestana '{SHEET_MFGPRO}'.")
    return excel_file


def read_charted_raw(excel_file):
    """Lee Charted saltando las filas de encabezado que no son tabla."""
    return pd.read_excel(excel_file, sheet_name=SHEET_CHARTED, skiprows=8, header=0)


def read_mfgpro_raw(excel_file):
    """Lee MFGPro saltando sus filas iniciales de reporte."""
    return pd.read_excel(excel_file, sheet_name=SHEET_MFGPRO, skiprows=3, header=0)


def load_components_df(path, sheet_name=SHEET_COMPONENTS):
    """Carga la BD de componentes que aporta unidad de medida por parte."""
    path = _resolve_excel_path(path)
    return pd.read_excel(path, sheet_name=sheet_name, header=0)
