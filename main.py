import sys

from compare import procesar_charted, procesar_mfgpro, comparar_bom
from export import guardar_resultado
from io_excel import (
    SHEET_CHARTED,
    SHEET_MFGPRO,
    load_excel_file,
    load_charted_file,
    load_mfgpro_file,
    read_charted_raw,
    read_mfgpro_raw,
    load_components_df,
)


DEFAULT_COMPONENTS_PATH = "data/bd_componentes.xls"


def calcular_resultado(
    ruta_charted,
    ruta_mfgpro=None,
    ruta_componentes=DEFAULT_COMPONENTS_PATH,
    hoja_componentes="Data",
):
    """Carga, limpia y compara los BOMs; devuelve datos y conteos para reportar."""
    if ruta_mfgpro:
        # Flujo nuevo: cada BOM viene en su archivo original.
        charted_file = load_charted_file(ruta_charted)
        mfgpro_file = load_mfgpro_file(ruta_mfgpro)
        try:
            df_charted_raw = read_charted_raw(charted_file)
            df_mfgpro_raw = read_mfgpro_raw(mfgpro_file)
        finally:
            charted_file.close()
            mfgpro_file.close()
    else:
        # Flujo viejo de pruebas: un solo Excel con ambas hojas ya pegadas.
        excel_file = load_excel_file(ruta_charted)
        try:
            df_charted_raw = read_charted_raw(excel_file, sheet_name=SHEET_CHARTED)
            df_mfgpro_raw = read_mfgpro_raw(excel_file, sheet_name=SHEET_MFGPRO)
        finally:
            excel_file.close()

    df_charted, model_cols_c, col_issue_map = procesar_charted(df_charted_raw)

    df_mfgpro, model_cols_m = procesar_mfgpro(df_mfgpro_raw)

    # La BD de componentes ayuda a completar unidades cuando el BOM no las trae.
    df_componentes = load_components_df(ruta_componentes, sheet_name=hoja_componentes)
    df_resultado = comparar_bom(
        df_charted,
        model_cols_c,
        df_mfgpro,
        model_cols_m,
        col_issue_map,
        df_components=df_componentes,
    )

    resumen = None
    if not df_resultado.empty:
        # Conteo simple para mostrar cuantos agrega/borra/modifica.
        resumen = df_resultado["_estado"].value_counts().to_dict()

    return (
        df_resultado,
        len(df_charted),
        len(model_cols_c),
        len(df_mfgpro),
        len(model_cols_m),
        resumen,
    )


def obtener_partes_afectadas(df_resultado):
    """Devuelve numeros de parte unicos que aparecen en las diferencias."""
    if df_resultado.empty or "Part Nbr." not in df_resultado.columns:
        return []

    partes = df_resultado["Part Nbr."].dropna().astype(str).str.strip()
    partes = partes[partes != ""]
    return list(dict.fromkeys(partes))


def proceso_calcular(
    ruta_charted,
    ruta_salida,
    ruta_mfgpro=None,
    ruta_componentes=DEFAULT_COMPONENTS_PATH,
    hoja_componentes="Data",
):
    """Version de consola: calcula diferencias y escribe el Excel final."""
    print("Cargando datos...")
    (
        df_resultado,
        charted_partes,
        charted_modelos,
        mfg_partes,
        mfg_modelos,
        resumen,
    ) = calcular_resultado(
        ruta_charted,
        ruta_mfgpro=ruta_mfgpro,
        ruta_componentes=ruta_componentes,
        hoja_componentes=hoja_componentes,
    )

    print(f"   Charted: {charted_partes} partes, {charted_modelos} modelos")
    print(f"   MFGPro:  {mfg_partes} partes, {mfg_modelos} modelos")

    if df_resultado.empty:
        print("No se encontraron diferencias entre los BOMs.")
    else:
        print(f"   -> {len(df_resultado)} diferencias: {resumen}")
        print("Numeros de parte afectados:")
        for parte in obtener_partes_afectadas(df_resultado):
            print(f"   - {parte}")

    print("Guardando resultado...")
    guardar_resultado(df_resultado, ruta_salida)
    print("Proceso completado.")


if __name__ == "__main__":
    # Con argumentos corre en consola; sin argumentos abre la interfaz grafica.
    if len(sys.argv) > 1:
        archivo_charted = sys.argv[1]
        archivo_mfgpro = sys.argv[2] if len(sys.argv) > 3 else None
        archivo_salida = sys.argv[3] if len(sys.argv) > 3 else (
            sys.argv[2] if len(sys.argv) > 2 else "Resultado_BOM.xlsx"
        )
        try:
            proceso_calcular(
                archivo_charted,
                archivo_salida,
                ruta_mfgpro=archivo_mfgpro,
            )
        except Exception as e:
            print(f"Error: {e}")
            raise
    else:
        from gui import run_app

        run_app()
