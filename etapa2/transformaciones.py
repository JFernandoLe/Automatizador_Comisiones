import pandas as pd


COLUMNAS_FINALES = [
    "ARCHIVO EXTRAIDO",
    "FECHA",
    "MES",
    "AGENTE",
    "NOMBREAGENTE",
    "INTCONCEPTO",
    "DESCCONCEPTO",
    "RAMO",
    "POLIZA",
    "PRIMA",
    "COMISION",
    "USUARIO",
    "INTCONSECUTIVO",
    "PROMOTORIA",
    "DESCPROMOTORIA",
    "TIPOPERSONA",
    "CP",
    "DESCPRESP",
    "PROMOTORIA_1",
    "DESCPROMOTORIA_1"
]

MESES_ES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE"
}


def transformar_saa(df):

    df = df.copy()

    df["FECHA"] = pd.to_datetime(
        df["FECHA"],
        format="%d-%b-%y",
        errors="coerce"
    )

    return df


def transformar_manuales(df):

    salida = pd.DataFrame()

    salida["FECHA"] = pd.to_datetime(
        df["Fecha"],
        errors="coerce"
    )

    salida["AGENTE"] = df["Agente"]
    salida["NOMBREAGENTE"] = df["Nombre"]

    salida["INTCONCEPTO"] = df["Cpto"]
    salida["DESCCONCEPTO"] = df["Desc Cpto"]

    salida["RAMO"] = pd.to_numeric(
        df["Ramo"],
        errors="coerce"
    )

    salida["POLIZA"] = df["Póliza"]

    salida["PRIMA"] = df["Prima"]
    salida["COMISION"] = df["Importe"]

    salida["USUARIO"] = df["Usuario"]

    salida["INTCONSECUTIVO"] = df["Consecutivo"]

    salida["PROMOTORIA"] = df["Promotoria"]

    salida["DESCPROMOTORIA"] = (
        df["Desc Promotoria"]
    )

    salida["TIPOPERSONA"] = (
        df["Tipo de Persona"]
    )

    salida["CP"] = df["CP"]

    salida["DESCPRESP"] = df["Desc CP"]

    salida["PROMOTORIA_1"] = (
        salida["PROMOTORIA"]
    )

    salida["DESCPROMOTORIA_1"] = (
        salida["DESCPROMOTORIA"]
    )

    return salida


def construir_vida(
    sap_vida,
    saa,
    manuales
):

    sap_vida = sap_vida.copy()
    sap_vida["FUENTE"] = "SAP"

    saa = transformar_saa(
        saa.copy()
    )

    saa = saa[
        saa["RAMO"] == 101
    ].copy()

    saa["FUENTE"] = "SAA"

    manuales = transformar_manuales(
        manuales.copy()
    )

    manuales = manuales[
        manuales["RAMO"] == 101
    ].copy()

    manuales["FUENTE"] = (
        "SAA MANUAL"
    )

    resultado = pd.concat(
        [
            sap_vida,
            saa,
            manuales
        ],
        ignore_index=True
    )
    resultado["AGENTE"] = (
    resultado["AGENTE"]
    .fillna("")
    .astype(str)
    )

    resultado["INTCONSECUTIVO"] = (
        resultado["INTCONSECUTIVO"]
        .fillna("")
        .astype(str)
    )

    resultado["POLIZA"] = (
        resultado["POLIZA"]
        .fillna("")
        .astype(str)
    )

    resultado["FECHA"] = pd.to_datetime(
    resultado["FECHA"],
    errors="coerce"
    )

    print(
        "\nFechas nulas VIDA:",
        resultado["FECHA"].isna().sum()
    )

    print(
        resultado.loc[
            resultado["FECHA"].isna(),
            ["FUENTE"]
        ].value_counts()
    )

    resultado["MES"] = (
        resultado["FECHA"]
        .dt.month
        .map(MESES_ES)
    )

    resultado["ARCHIVO EXTRAIDO"] = (
        resultado["FUENTE"]
    )
    resultado["INTCONCEPTO"] = pd.to_numeric(
    resultado["INTCONCEPTO"],
    errors="coerce"
    )

    resultado["RAMO"] = pd.to_numeric(
        resultado["RAMO"],
        errors="coerce"
    )

    resultado["PRIMA"] = pd.to_numeric(
        resultado["PRIMA"],
        errors="coerce"
    )

    resultado["COMISION"] = pd.to_numeric(
        resultado["COMISION"],
        errors="coerce"
    )

    resultado["CP"] = pd.to_numeric(
        resultado["CP"],
        errors="coerce"
    )
    return resultado[
        COLUMNAS_FINALES
    ]


def construir_gmm(
    sap_gmm,
    saa,
    manuales
):

    sap_gmm = sap_gmm.copy()
    sap_gmm["FUENTE"] = "SAP"

    saa = transformar_saa(
        saa.copy()
    )

    saa = saa[
        saa["RAMO"] == 300
    ].copy()

    saa["FUENTE"] = "SAA"

    manuales = transformar_manuales(
        manuales.copy()
    )

    manuales = manuales[
        manuales["RAMO"] == 300
    ].copy()

    manuales["FUENTE"] = (
        "SAA MANUAL"
    )

    resultado = pd.concat(
        [
            sap_gmm,
            saa,
            manuales
        ],
        ignore_index=True
    )
    resultado["AGENTE"] = (
    resultado["AGENTE"]
    .fillna("")
    .astype(str)
    )

    resultado["INTCONSECUTIVO"] = (
        resultado["INTCONSECUTIVO"]
        .fillna("")
        .astype(str)
    )

    resultado["POLIZA"] = (
        resultado["POLIZA"]
        .fillna("")
        .astype(str)
    )


    resultado["FECHA"] = pd.to_datetime(
    resultado["FECHA"],
    errors="coerce"
    )

    print(
        "\nFechas nulas GMM:",
        resultado["FECHA"].isna().sum()
    )
    print(
        resultado.loc[
            resultado["FECHA"].isna(),
            ["FUENTE"]
        ].value_counts()
    )

    resultado["MES"] = (
        resultado["FECHA"]
        .dt.month
        .map(MESES_ES)
    )

    resultado["ARCHIVO EXTRAIDO"] = (
        resultado["FUENTE"]
    )
    resultado["INTCONCEPTO"] = pd.to_numeric(
    resultado["INTCONCEPTO"],
    errors="coerce"
    )

    resultado["RAMO"] = pd.to_numeric(
        resultado["RAMO"],
        errors="coerce"
    )

    resultado["PRIMA"] = pd.to_numeric(
        resultado["PRIMA"],
        errors="coerce"
    )

    resultado["COMISION"] = pd.to_numeric(
        resultado["COMISION"],
        errors="coerce"
    )

    resultado["CP"] = pd.to_numeric(
        resultado["CP"],
        errors="coerce"
    )
    return resultado[
        COLUMNAS_FINALES
    ]