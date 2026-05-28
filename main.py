import sys

from compare import procesar_charted, procesar_mfgpro, comparar_bom
from export import guardar_resultado
from io_excel import (
    load_excel_file,
    read_charted_raw,
    read_mfgpro_raw,
    load_components_df,
)


def calcular_resultado(
    ruta_archivo,
    ruta_componentes="data/bd_componentes.xlsx",
    hoja_componentes="Data",
):
    excel_file = load_excel_file(ruta_archivo)

    df_charted_raw = read_charted_raw(excel_file)
    df_charted, model_cols_c, col_issue_map = procesar_charted(df_charted_raw)

    df_mfgpro_raw = read_mfgpro_raw(excel_file)
    df_mfgpro, model_cols_m = procesar_mfgpro(df_mfgpro_raw)

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
        resumen = df_resultado["_estado"].value_counts().to_dict()

    return df_resultado, len(df_charted), len(model_cols_c), len(df_mfgpro), len(model_cols_m), resumen


def proceso_calcular(
    ruta_archivo,
    ruta_salida,
    ruta_componentes="data/bd_componentes.xlsx",
    hoja_componentes="Data",
):
    print("Cargando datos...")
    df_resultado, charted_partes, charted_modelos, mfg_partes, mfg_modelos, resumen = calcular_resultado(
        ruta_archivo,
        ruta_componentes=ruta_componentes,
        hoja_componentes=hoja_componentes,
    )

    print(f"   Charted: {charted_partes} partes, {charted_modelos} modelos")
    print(f"   MFGPro:  {mfg_partes} partes, {mfg_modelos} modelos")

    if df_resultado.empty:
        print("No se encontraron diferencias entre los BOMs.")
    else:
        print(f"   -> {len(df_resultado)} diferencias: {resumen}")

    print("Guardando resultado...")
    guardar_resultado(df_resultado, ruta_salida)
    print("Proceso completado.")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        archivo_entrada = sys.argv[1]
        archivo_salida = sys.argv[2] if len(sys.argv) > 2 else "Resultado_BOM.xlsx"
        try:
            proceso_calcular(archivo_entrada, archivo_salida)
        except Exception as e:
            print(f"Error: {e}")
            raise
    else:
        from gui import run_app

        run_app()
