import numpy as np
import pandas as pd


CONVERSION = {
    "FT ": lambda x: round(x / 304.8, 0),
    "FT": lambda x: round(x / 304.8, 3),
    "MT": lambda x: round(x / 1000, 3),
    "M": lambda x: round(x / 1000, 3),
    "LB": lambda x: round(x / 453.5924, 5),
    "KG": lambda x: round(x / 1000, 5),
    "EA": lambda x: round(x, 5),
}

DECIMALES = {
    "FT ": 0,
    "FT": 3,
    "MT": 3,
    "M": 3,
    "LB": 5,
    "KG": 5,
    "EA": 5,
}

UMBRAL = {
    "FT ": 1 / 304.8,
    "FT": 1 / 304.8,
    "MT": 1 / 1000,
    "M": 1 / 1000,
    "LB": 1e-9,
    "KG": 1e-9,
    "EA": 1.0,
}


def convertir(qty, unit):
    if pd.isna(qty):
        return 0.0
    qty = float(qty)
    if pd.isna(unit):
        return qty
    fn = CONVERSION.get(str(unit).strip())
    return fn(qty) if fn else qty


def redondear_old(qty, unit):
    decimales = DECIMALES.get(str(unit).strip(), 5)
    return round(float(qty), decimales)


def procesar_charted(df):
    df.dropna(how="all", inplace=True)

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
    df.dropna(how="all", inplace=True)

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
    modelos = [c for c in model_cols_c if c in model_cols_m]

    idx_charted = df_charted.set_index("EPN")
    idx_mfgpro = df_mfgpro.set_index("Item")

    todas_partes = set(idx_charted.index) | set(idx_mfgpro.index)

    filas = []

    uom_lookup = _build_uom_lookup(df_components, part_col, uom_col)

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
            qty_raw_new = idx_charted.loc[part, modelo] if en_charted else np.nan
            qty_new_conv = convertir(qty_raw_new, unit)

            qty_raw_old = idx_mfgpro.loc[part, modelo] if en_mfgpro else 0.0
            qty_old = float(qty_raw_old) if not pd.isna(qty_raw_old) else 0.0

            decimales = DECIMALES.get(str(unit).strip(), 5)
            diferencia = qty_new_conv - qty_old

            umbral = UMBRAL.get(str(unit).strip(), 1e-9)
            if abs(diferencia) < umbral:
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
