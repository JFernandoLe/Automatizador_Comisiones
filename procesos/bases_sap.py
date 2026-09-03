from conversores.preparar import leer_como_automatizador_original, preparar_excel_sap
from etapa1.procesador import procesar_dataframe
from servicios.estado import en_rango
from servicios.excel import guardar_excel_dividido


def generar_bases_sap(vida_file, gmm_file, hojas_vida, hojas_gmm, actualizar_estado):
    if not vida_file:
        raise ValueError("Debe seleccionar uno o más archivos VIDA.")
    if not gmm_file:
        raise ValueError("Debe seleccionar uno o más archivos GMM.")

    actualizar_estado("Preparando Base VIDA...", 2)
    vida_convertido = preparar_excel_sap(
        vida_file,
        hojas_seleccionadas=hojas_vida,
        prefijo="vida_",
        actualizar_estado=en_rango(actualizar_estado, 2, 30),
    )
    actualizar_estado("Leyendo Base VIDA...", 32)
    df_vida = leer_como_automatizador_original(
        vida_file, vida_convertido, hojas_vida
    )
    actualizar_estado("Procesando Base VIDA...", 36)
    df_vida = procesar_dataframe(df_vida, "VIDA")

    actualizar_estado("Preparando Base GMM...", 50)
    gmm_convertido = preparar_excel_sap(
        gmm_file,
        hojas_seleccionadas=hojas_gmm,
        prefijo="gmm_",
        actualizar_estado=en_rango(actualizar_estado, 50, 76),
    )
    actualizar_estado("Leyendo Base GMM...", 78)
    df_gmm = leer_como_automatizador_original(
        gmm_file, gmm_convertido, hojas_gmm
    )
    actualizar_estado("Procesando Base GMM...", 82)
    df_gmm = procesar_dataframe(df_gmm, "GMM")

    actualizar_estado("Generando archivos de salida...", 92)
    df_vida.to_parquet("vida_para_reporte.parquet", index=False)
    df_gmm.to_parquet("gmm_para_reporte.parquet", index=False)
    guardar_excel_dividido(df_vida, "vida_para_reporte.xlsx")
    guardar_excel_dividido(df_gmm, "gmm_para_reporte.xlsx")

    actualizar_estado("Bases SAP generadas", 100)
    return df_vida, df_gmm
