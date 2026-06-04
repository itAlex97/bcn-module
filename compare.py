import numpy as np
import pandas as pd


# Cada unidad vive aqui una sola vez: divisor para convertir, decimales
# para mostrar y umbral minimo para ignorar diferencias de redondeo.
UNIT_RULES = {
    "FT": {"divisor": 304.8, "decimales": 3, "umbral": 1 / 304.8},
    "MT": {"divisor": 1000, "decimales": 3, "umbral": 1 / 1000},
    "M": {"divisor": 1000, "decimales": 3, "umbral": 1 / 1000},
    "LB": {"divisor": 453.5924, "decimales": 5, "umbral": 1e-9},
    "KG": {"divisor": 1000, "decimales": 5, "umbral": 1e-9},
    "EA": {"divisor": 1, "decimales": 5, "umbral": 1.0},
}

DEFAULT_DECIMALES = 5
DEFAULT_UMBRAL = 1e-9

# Alias derivados para no repetir configuracion y conservar nombres conocidos.
DECIMALES = {unit: rule["decimales"] for unit, rule in UNIT_RULES.items()}
UMBRAL = {unit: rule["umbral"] for unit, rule in UNIT_RULES.items()}


def convertir(qty, unit):
    """Convierte la cantidad nueva segun la unidad usada por MFGPro/componentes."""
    if pd.isna(qty):
        return 0.0
    qty = float(qty)
    rule = UNIT_RULES.get(_unit_key(unit))
    if not rule:
        return qty
    return round(qty / rule["divisor"], rule["decimales"])


def redondear_old(qty, unit):
    """Redondea cantidades existentes con la precision esperada por unidad."""
    return round(float(qty), _decimales(unit))


def _unit_key(unit):
    """Normaliza la unidad para poder buscarla en UNIT_RULES."""
    return "" if pd.isna(unit) else str(unit).strip()


def _decimales(unit):
    """Devuelve cuantos decimales se usan para una unidad."""
    return DECIMALES.get(_unit_key(unit), DEFAULT_DECIMALES)


def _umbral(unit):
    """Devuelve la tolerancia minima antes de reportar una diferencia."""
    return UMBRAL.get(_unit_key(unit), DEFAULT_UMBRAL)


def procesar_charted(df):
    """Limpia Charted BOM y detecta las columnas que representan modelos."""
    df.dropna(how="all", inplace=True)

    # Se renombran solo las columnas fijas que vienen por posicion.
    df.rename(
        columns={
            df.columns[0]: "EPN",
            df.columns[5]: "Part Description",
            df.columns[6]: "UOM MM/G",
        },
        inplace=True,
    )

    if "TOTAL" in df.columns:
        df.drop(columns=["TOTAL"], inplace=True)

    # Los encabezados de modelos pueden traer issue/revision; la base es
    # lo que se compara contra MFGPro y el mapa conserva el texto original.
    col_issue_map = {}
    nuevos_nombres = {}
    for col in df.columns:
        col_str = str(col)
        if col_str not in (
            "EPN",
            "Type",
            "Group",
            "Mat Code",
            "CPN",
            "Part Description",
            "UOM MM/G",
        ):
            base = col_str.split(" ")[0]
            nuevos_nombres[col] = base
            col_issue_map[base] = col_str
    df.rename(columns=nuevos_nombres, inplace=True)

    # Todo lo que no es metadata de la parte se trata como modelo.
    model_cols = [
        c
        for c in df.columns
        if c not in (
            "EPN",
            "Type",
            "Group",
            "Mat Code",
            "CPN",
            "Part Description",
            "UOM MM/G",
        )
    ]
    for col in model_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df[df["EPN"].notna() & (df["EPN"].astype(str).str.strip() != "")]
    df.reset_index(drop=True, inplace=True)
    return df, model_cols, col_issue_map


def procesar_mfgpro(df):
    """Limpia MFGPro y deja sus cantidades listas para comparar."""
    df.dropna(how="all", inplace=True)

    # MFGPro trae las primeras columnas fijas y despues vienen los modelos.
    df.rename(
        columns={
            df.columns[0]: "Item",
            df.columns[1]: "Description",
            df.columns[2]: "Unit",
        },
        inplace=True,
    )

    model_cols = list(df.columns[3:])
    for col in model_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df = df[df["Item"].notna() & (df["Item"].astype(str).str.strip() != "")]
    df.reset_index(drop=True, inplace=True)
    return df, model_cols


def _build_uom_lookup(df_components, part_col, uom_col):
    """Crea un diccionario parte -> unidad desde la BD de componentes."""
    if df_components is None:
        return {}
    if part_col not in df_components.columns or uom_col not in df_components.columns:
        return {}
    lookup = {}
    for _, row in df_components.iterrows():
        part = str(row.get(part_col, "")).strip()
        uom = str(row.get(uom_col, "")).strip()
        if part and uom:
            lookup[part] = uom
    return lookup


def comparar_bom(
    df_charted,
    model_cols_c,
    df_mfgpro,
    model_cols_m,
    col_issue_map,
    df_components=None,
    part_col="Item Number",
    uom_col="Unit of Measure",
):
    """Compara Charted contra MFGPro y devuelve solo las diferencias."""
    modelos = [c for c in model_cols_c if c in model_cols_m]

    idx_charted = df_charted.set_index("EPN")
    idx_mfgpro = df_mfgpro.set_index("Item")

    todas_partes = set(idx_charted.index) | set(idx_mfgpro.index)

    filas = []

    uom_lookup = _build_uom_lookup(df_components, part_col, uom_col)

    # Se recorren todas las partes existentes en cualquiera de los dos BOMs.
    for part in todas_partes:
        part_str = str(part).strip()
        if part_str.lower() == "eliminar":
            continue

        en_charted = part in idx_charted.index
        en_mfgpro = part in idx_mfgpro.index

        desc = ""
        unit = ""
        if en_mfgpro:
            row_m = idx_mfgpro.loc[part]
            unit = str(row_m["Unit"]).strip() if pd.notna(row_m["Unit"]) else ""
            desc = str(row_m["Description"]) if pd.notna(row_m["Description"]) else ""
        if not desc and en_charted:
            row_c = idx_charted.loc[part]
            desc = (
                str(row_c["Part Description"])
                if pd.notna(row_c["Part Description"])
                else ""
            )
        if not unit:
            unit = uom_lookup.get(part_str, "")

        for modelo in modelos:
            # Charted se convierte a la unidad esperada antes de comparar.
            qty_raw_new = idx_charted.loc[part, modelo] if en_charted else np.nan
            qty_new_conv = convertir(qty_raw_new, unit)

            qty_raw_old = idx_mfgpro.loc[part, modelo] if en_mfgpro else 0.0
            qty_old = float(qty_raw_old) if not pd.isna(qty_raw_old) else 0.0

            decimales = _decimales(unit)
            diferencia = qty_new_conv - qty_old

            if abs(diferencia) < _umbral(unit):
                continue

            if not en_mfgpro or qty_old == 0:
                estado = "Se agrega"
            elif not en_charted or qty_new_conv == 0:
                estado = "Se Borra"
            else:
                estado = "Se Modifica"

            filas.append(
                {
                    "Part Nbr.": part_str,
                    "Description": desc,
                    "UOM": unit,
                    "Qty New": round(qty_new_conv, decimales),
                    "Qty Old": round(qty_old, decimales),
                    "Model": modelo,
                    "_modelo_issue": col_issue_map.get(modelo, modelo),
                    "_estado": estado,
                }
            )

    df = pd.DataFrame(filas)
    return df
