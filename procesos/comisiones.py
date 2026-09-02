from conversores.preparar import preparar_manuales, preparar_saa
from etapa2.procesador import generar_vida, generar_gmm
from servicios.excel import guardar_excel_dividido


def generar_comisiones(
    saa_file,
    manuales_file,
    ejecutar_vida,
    ejecutar_gmm,
    actualizar_estado,
    hojas_saa=None,
    hojas_manuales=None,
):
    if not saa_file:
        raise ValueError("Debe seleccionar el archivo SAA.")
    if not manuales_file:
        raise ValueError("Debe seleccionar uno o más archivos de Acumulado Comisiones.")
    if not ejecutar_vida and not ejecutar_gmm:
        raise ValueError("Debe seleccionar VIDA o GMM.")

    actualizar_estado("Preparando SAA...", 4)
    saa_listo = preparar_saa(saa_file, hojas_saa)
    actualizar_estado("Preparando Acumulado Comisiones...", 7)
    manuales_listo = preparar_manuales(
        manuales_file,
        hojas_manuales,
        actualizar_estado=actualizar_estado,
    )

    resultados = {}

    if ejecutar_vida:
        actualizar_estado("Generando comisiones VIDA...", 10)
        df_vida = generar_vida(saa_listo, manuales_listo)
        df_vida.to_parquet("vida_comisiones.parquet", index=False)
        guardar_excel_dividido(df_vida, "vida_comisiones.xlsx")
        resultados["vida"] = df_vida

    if ejecutar_gmm:
        actualizar_estado("Generando comisiones GMM...", 60)
        df_gmm = generar_gmm(saa_listo, manuales_listo)
        df_gmm.to_parquet("gmm_comisiones.parquet", index=False)
        guardar_excel_dividido(df_gmm, "gmm_comisiones.xlsx")
        resultados["gmm"] = df_gmm

    actualizar_estado("Comisiones generadas", 100)
    return resultados
