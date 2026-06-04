import os
import warnings
import pandas as pd


# Nombres fijos de hojas esperadas en los archivos Excel.
SHEET_CHARTED = "Charted_BOM_CapXC"
SHEET_MFGPRO = "MFGPro_CAPH"
SHEET_CHARTED_ORIGINAL = "Detail"
SHEET_MFGPRO_ORIGINAL = "Sheet1"
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


def _find_sheet_name(excel_file, expected_name):
    """Busca una hoja por nombre sin depender de mayusculas/minusculas."""
    sheet_names = excel_file.sheet_names
    if expected_name in sheet_names:
        return expected_name

    expected_lower = expected_name.lower()
    for sheet_name in sheet_names:
        if str(sheet_name).lower() == expected_lower:
            return sheet_name

    raise ValueError(f"No se encontro la pestana '{expected_name}'.")


def load_excel_file(path):
    """Abre el Excel viejo de pruebas y valida sus dos hojas internas."""
    path = _resolve_excel_path(path)
    excel_file = pd.ExcelFile(path)
    _find_sheet_name(excel_file, SHEET_CHARTED)
    _find_sheet_name(excel_file, SHEET_MFGPRO)
    return excel_file


def load_charted_file(path):
    """Abre el archivo original de Charted y valida la hoja detail."""
    path = _resolve_excel_path(path)
    excel_file = pd.ExcelFile(path)
    _find_sheet_name(excel_file, SHEET_CHARTED_ORIGINAL)
    return excel_file


def load_mfgpro_file(path):
    """Abre el archivo original de MFGPro y valida la hoja Sheet1."""
    path = _resolve_excel_path(path)
    excel_file = pd.ExcelFile(path)
    _find_sheet_name(excel_file, SHEET_MFGPRO_ORIGINAL)
    return excel_file


def read_charted_raw(excel_file, sheet_name=SHEET_CHARTED_ORIGINAL):
    """Lee Charted saltando las filas de encabezado que no son tabla."""
    sheet_name = _find_sheet_name(excel_file, sheet_name)
    return _read_excel_sheet(excel_file, sheet_name=sheet_name, skiprows=8)


def read_mfgpro_raw(excel_file, sheet_name=SHEET_MFGPRO_ORIGINAL):
    """Lee MFGPro saltando sus filas iniciales de reporte."""
    sheet_name = _find_sheet_name(excel_file, sheet_name)
    return _read_excel_sheet(excel_file, sheet_name=sheet_name, skiprows=3)


def load_components_df(path, sheet_name=SHEET_COMPONENTS):
    """Carga la BD de componentes que aporta unidad de medida por parte."""
    path = _resolve_excel_path(path)
    return _read_excel_sheet(path, sheet_name=sheet_name, skiprows=0)


def _read_excel_sheet(excel_source, sheet_name, skiprows):
    """Lee una hoja y oculta solo el warning de estilo default de openpyxl."""
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="Workbook contains no default style, apply openpyxl's default",
            category=UserWarning,
            module="openpyxl.styles.stylesheet",
        )
        return pd.read_excel(
            excel_source,
            sheet_name=sheet_name,
            skiprows=skiprows,
            header=0,
        )
