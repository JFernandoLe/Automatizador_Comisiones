from etapa2.procesador import generar_vida, generar_gmm
from servicios.excel import guardar_excel_dividido


def generar_comisiones(
    saa_file,
    manuales_file,
    ejecutar_vida,
    ejecutar_gmm,
    actualizar_estado,
):
    if not saa_file:
        raise ValueError("Debe seleccionar el archivo SAA.")
    if not manuales_file:
        raise ValueError("Debe seleccionar el archivo de Manuales.")
    if not ejecutar_vida and not ejecutar_gmm:
        raise ValueError("Debe seleccionar VIDA o GMM.")

    resultados = {}

    if ejecutar_vida:
        actualizar_estado("Generando comisiones VIDA...", 10)
        df_vida = generar_vida(saa_file, manuales_file)
        df_vida.to_parquet("vida_comisiones.parquet", index=False)
        guardar_excel_dividido(df_vida, "vida_comisiones.xlsx")
        resultados["vida"] = df_vida

    if ejecutar_gmm:
        actualizar_estado("Generando comisiones GMM...", 60)
        df_gmm = generar_gmm(saa_file, manuales_file)
        df_gmm.to_parquet("gmm_comisiones.parquet", index=False)
        guardar_excel_dividido(df_gmm, "gmm_comisiones.xlsx")
        resultados["gmm"] = df_gmm

    actualizar_estado("Comisiones generadas", 100)
    return resultados
