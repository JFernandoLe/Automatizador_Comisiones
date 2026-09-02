import pandas as pd

from etapa4.transformaciones import construir_vida_etapa4, construir_gmm_etapa4

HOJAS_CATALOGOS = {
    "PFPM": ["AGENTE_ORIGINAL", "AGENTE_REPORTERIA"],
    "CONCEPTOS_VIDA": ["INTCONCEPTO", "TIPO_COMISION", "COMO_PROMOTOR_AGENTE"],
    "CONCEPTOS_GMM": ["INTCONCEPTO", "TIPO_COMISION", "COMO_PROMOTOR_AGENTE"],
}


def _normalizar_columnas(df):
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]
    return df


def _validar_columnas(df, hoja, requeridas):
    faltantes = [col for col in requeridas if col not in df.columns]
    if faltantes:
        raise ValueError(
            f"La hoja '{hoja}' no contiene las columnas obligatorias: "
            + ", ".join(faltantes)
        )


def cargar_catalogos(ruta_catalogos):
    if not ruta_catalogos:
        raise ValueError("Debe seleccionar el archivo único de catálogos.")

    excel = pd.ExcelFile(ruta_catalogos, engine="openpyxl")
    hojas_disponibles = {str(hoja).strip().upper(): hoja for hoja in excel.sheet_names}
    faltantes = [hoja for hoja in HOJAS_CATALOGOS if hoja not in hojas_disponibles]
    if faltantes:
        raise ValueError(
            "El archivo de catálogos no contiene las hojas obligatorias: "
            + ", ".join(faltantes)
        )

    dataframes = {}
    for hoja, columnas in HOJAS_CATALOGOS.items():
        df = pd.read_excel(
            ruta_catalogos,
            sheet_name=hojas_disponibles[hoja],
            engine="openpyxl",
        )
        df = _normalizar_columnas(df)
        _validar_columnas(df, hoja, columnas)
        dataframes[hoja] = df

    return dataframes["PFPM"], dataframes["CONCEPTOS_VIDA"], dataframes["CONCEPTOS_GMM"]


def generar_vida_etapa4(ruta_vlsp, hoja_vlsp, ruta_tipo, hoja_tipo, ruta_catalogos):
    print("\n" + "=" * 80)
    print("ETAPA 4 VIDA")
    print("=" * 80)

    df_vida = pd.read_parquet("vida_comisiones.parquet")
    df_vlsp = pd.read_excel(ruta_vlsp, sheet_name=hoja_vlsp, engine="openpyxl")
    df_tipo = pd.read_excel(
        ruta_tipo, sheet_name=hoja_tipo, header=2, engine="openpyxl"
    )
    df_pfpm, df_conceptos_vida, _ = cargar_catalogos(ruta_catalogos)

    resultado = construir_vida_etapa4(
        df_vida, df_vlsp, df_tipo, df_pfpm, df_conceptos_vida
    )
    print(f"Resultado VIDA: {len(resultado):,}")
    return resultado


def generar_gmm_etapa4(ruta_catalogos):
    print("\n" + "=" * 80)
    print("ETAPA 4 GMM")
    print("=" * 80)

    df_gmm = pd.read_parquet("gmm_comisiones.parquet")
    df_pfpm, _, df_conceptos_gmm = cargar_catalogos(ruta_catalogos)

    resultado = construir_gmm_etapa4(df_gmm, df_pfpm, df_conceptos_gmm)
    print(f"Resultado GMM: {len(resultado):,}")
    return resultado
