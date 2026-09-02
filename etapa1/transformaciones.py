import pandas as pd

from etapa1.catalogos import (
    INTCONCEPTO_MAP,
    REPORTE_FINAL_MAP
)


def limpiar_columnas(df):

    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.strip()
    )

    return df


def obtener_columna(df, nombre):

    for col in df.columns:

        if nombre.upper() in col.upper():

            return col

    raise Exception(
        f"No se encontró la columna '{nombre}'.\n"
        f"Columnas encontradas:\n{list(df.columns)}"
    )


def procesar_vida(df):

    print("\nProcesando VIDA...")

    df = limpiar_columnas(df)

    concepto_col = obtener_columna(df, "Concepto")

    df["INTCONCEPTO"] = (
        df[concepto_col]
        .astype(str)
        .str.strip()
        .map(INTCONCEPTO_MAP)
    )

    df["REPORTE FINAL DE COMISIONES"] = (
        df[concepto_col]
        .astype(str)
        .str.strip()
        .map(REPORTE_FINAL_MAP)
    )

    registros_antes = len(df)

    df = df[
        df["REPORTE FINAL DE COMISIONES"] == "SI"
    ].copy()

    print(
        f"Filtrados VIDA: "
        f"{registros_antes:,} -> {len(df):,}"
    )

    df["RAMO"] = df["Ramo"].apply(
        lambda x:
        101
        if str(x).strip().upper() in ["VIDA", "101"]
        else "VERIFICAR"
    )

    df["INTCONSECUTIVO"] = ""

    df["PROMOTORIA_1"] = df["Promotoria"]

    df["DESCPROMOTORIA_1"] = (
        df["Descripción Promotoria"]
    )

    salida = df[
        [
            "Fecha",
            "Agente",
            "Nombre Agente",
            "INTCONCEPTO",
            concepto_col,
            "RAMO",
            "Poliza",
            "Prima",
            "Comisión",
            "Como Promotor / Agente",
            "INTCONSECUTIVO",
            "Promotoria",
            "Descripción Promotoria",
            "Tipo de Persona",
            "CP",
            "Tipo de Figura",
            "PROMOTORIA_1",
            "DESCPROMOTORIA_1",
        ]
    ].copy()

    salida.columns = [
        "FECHA",
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
        "DESCPROMOTORIA_1",
    ]

    salida["FECHA"] = pd.to_datetime(
        salida["FECHA"],
        errors="coerce"
    )

    return salida

def procesar_gmm(df):

    print("\nProcesando GMM...")

    df = limpiar_columnas(df)

    concepto_col = obtener_columna(df, "Concepto")

    subramo_col = obtener_columna(df, "Subramo")

    df["INTCONCEPTO"] = (
        df[concepto_col]
        .astype(str)
        .str.strip()
        .map(INTCONCEPTO_MAP)
    )

    df["REPORTE FINAL DE COMISIONES"] = (
        df[concepto_col]
        .astype(str)
        .str.strip()
        .map(REPORTE_FINAL_MAP)
    )

    df[subramo_col] = pd.to_numeric(
        df[subramo_col],
        errors="coerce"
    )

    registros_antes = len(df)

    df = df[
        df["REPORTE FINAL DE COMISIONES"] == "SI"
    ].copy()

    print(
        f"Filtro Reporte Final GMM: "
        f"{registros_antes:,} -> {len(df):,}"
    )

    df["RAMO"] = df["Ramo"].apply(
        lambda x:
        300
        if str(x).strip().upper() in ["GMM", "300"]
        else "VERIFICA"
    )

    df["INTCONSECUTIVO"] = ""

    df["PROMOTORIA_1"] = df["Promotoria"]

    df["DESCPROMOTORIA_1"] = (
        df["Descripción Promotoria"]
    )

    salida = df[
        [
            "Fecha",
            "Agente",
            "Nombre Agente",
            "INTCONCEPTO",
            concepto_col,
            "RAMO",
            "Poliza",
            "Prima",
            "Comisión",
            "Como Promotor / Agente",
            "INTCONSECUTIVO",
            "Promotoria",
            "Descripción Promotoria",
            "Tipo de Persona",
            "CP",
            "Descripción",
            "PROMOTORIA_1",
            "DESCPROMOTORIA_1",
        ]
    ].copy()

    salida.columns = [
        "FECHA",
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
        "DESCPROMOTORIA_1",
    ]

    salida["FECHA"] = pd.to_datetime(
        salida["FECHA"],
        errors="coerce"
    )

    return salida