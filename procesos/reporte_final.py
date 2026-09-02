from etapa4.procesador import generar_vida_etapa4, generar_gmm_etapa4
from servicios.excel import guardar_excel_dividido


def generar_reporte_final(
    vlsp_file,
    hoja_vlsp,
    tipo_file,
    hoja_tipo,
    catalogos_file,
    actualizar_estado,
):
    if not vlsp_file:
        raise ValueError("Debe seleccionar el archivo VLSP.")
    if not tipo_file:
        raise ValueError("Debe seleccionar el archivo Tipo.")
    if not catalogos_file:
        raise ValueError("Debe seleccionar el archivo único de catálogos.")
    if not hoja_vlsp:
        raise ValueError("Debe seleccionar una hoja VLSP.")
    if not hoja_tipo:
        raise ValueError("Debe seleccionar una hoja para el catálogo Tipo.")

    actualizar_estado("Generando reporte VIDA...", 25)
    df_vida = generar_vida_etapa4(
        vlsp_file, hoja_vlsp, tipo_file, hoja_tipo, catalogos_file
    )
    df_vida.to_parquet("Reporte_VIDA_Final.parquet", index=False)
    guardar_excel_dividido(df_vida, "Reporte_VIDA_Final.xlsx")

    actualizar_estado("Generando reporte GMM...", 75)
    df_gmm = generar_gmm_etapa4(catalogos_file)
    df_gmm.to_parquet("Reporte_GMM_Final.parquet", index=False)
    guardar_excel_dividido(df_gmm, "Reporte_GMM_Final.xlsx")

    actualizar_estado("Reporte final generado", 100)
    return df_vida, df_gmm
