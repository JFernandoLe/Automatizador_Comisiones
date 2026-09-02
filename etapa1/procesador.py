import pandas as pd

from etapa1.transformaciones import (
    procesar_vida,
    procesar_gmm
)


def procesar_dataframe(
    df,
    tipo
):

    print("=" * 80)
    print(f"PROCESANDO {tipo}")
    print("=" * 80)

    print(
        f"Registros recibidos: "
        f"{len(df):,}"
    )

    if tipo == "VIDA":

        df = procesar_vida(df)

    elif tipo == "GMM":

        df = procesar_gmm(df)

    else:

        raise Exception(
            f"Tipo no soportado: {tipo}"
        )

    print(
        f"Registros finales: "
        f"{len(df):,}"
    )

    return df


def procesar_archivo(
    archivo_origen,
    archivo_destino,
    tipo
):

    print("\n" + "=" * 80)
    print(
        f"PROCESANDO ARCHIVO: "
        f"{archivo_origen}"
    )
    print("=" * 80)

    df = pd.read_excel(
        archivo_origen,
        engine="openpyxl"
    )

    registros_originales = len(df)

    print(
        f"Registros originales: "
        f"{registros_originales:,}"
    )

    if tipo == "VIDA":

        df = procesar_vida(df)

    elif tipo == "GMM":

        df = procesar_gmm(df)

    else:

        raise Exception(
            f"Tipo no soportado: {tipo}"
        )

    registros_finales = len(df)

    print(
        f"Registros finales: "
        f"{registros_finales:,}"
    )

    try:

        prima_total = df["PRIMA"].sum()

        print(
            f"Prima total: "
            f"{prima_total:,.2f}"
        )

    except:

        pass

    try:

        comision_total = df["COMISION"].sum()

        print(
            f"Comisión total: "
            f"{comision_total:,.2f}"
        )

    except:

        pass

    df.to_excel(
        archivo_destino,
        index=False
    )

    print(
        f"Archivo generado: "
        f"{archivo_destino}"
    )

    return df