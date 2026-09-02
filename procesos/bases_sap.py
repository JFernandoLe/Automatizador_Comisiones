from conversores.preparar import preparar_excel_sap
from etapa1.procesador import procesar_dataframe
from servicios.excel import (
    guardar_excel_dividido,
    leer_hojas_seleccionadas,
    obtener_hojas,
)


def generar_bases_sap(vida_file, gmm_file, hojas_vida, hojas_gmm, actualizar_estado):
    if not vida_file:
        raise ValueError("Debe seleccionar uno o más archivos VIDA.")
    if not gmm_file:
        raise ValueError("Debe seleccionar uno o más archivos GMM.")

    actualizar_estado("Convirtiendo y consolidando VIDA...", 5)
    vida_convertido = preparar_excel_sap(
        vida_file,
        hojas_vida,
        prefijo="vida_",
        actualizar_estado=actualizar_estado,
    )
    actualizar_estado("Leyendo VIDA...", 10)
    df_vida = leer_hojas_seleccionadas(vida_convertido, obtener_hojas(vida_convertido))
    actualizar_estado("Procesando VIDA...", 25)
    df_vida = procesar_dataframe(df_vida, "VIDA")

    actualizar_estado("Convirtiendo y consolidando GMM...", 50)
    gmm_convertido = preparar_excel_sap(
        gmm_file,
        hojas_gmm,
        prefijo="gmm_",
        actualizar_estado=actualizar_estado,
    )
    actualizar_estado("Leyendo GMM...", 55)
    df_gmm = leer_hojas_seleccionadas(gmm_convertido, obtener_hojas(gmm_convertido))
    actualizar_estado("Procesando GMM...", 70)
    df_gmm = procesar_dataframe(df_gmm, "GMM")

    actualizar_estado("Guardando archivos...", 85)
    df_vida.to_parquet("vida_para_reporte.parquet", index=False)
    df_gmm.to_parquet("gmm_para_reporte.parquet", index=False)
    guardar_excel_dividido(df_vida, "vida_para_reporte.xlsx")
    guardar_excel_dividido(df_gmm, "gmm_para_reporte.xlsx")

    actualizar_estado("Bases SAP generadas", 100)
    return df_vida, df_gmm
