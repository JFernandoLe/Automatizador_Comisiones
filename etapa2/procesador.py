import pandas as pd

from etapa2.transformaciones import (
    construir_vida,
    construir_gmm
)


def generar_vida(
    ruta_saa,
    ruta_manuales
):

    print("\nCargando SAP VIDA...")

    sap_vida = pd.read_parquet(
        "vida_para_reporte.parquet"
    )

    print(
        f"Registros SAP VIDA: "
        f"{len(sap_vida):,}"
    )

    print("\nCargando SAA...")

    saa = pd.read_excel(
        ruta_saa,
        engine="openpyxl"
    )

    print(
        f"Registros SAA: "
        f"{len(saa):,}"
    )

    print("\nCargando MANUALES...")

    manuales = pd.read_excel(
        ruta_manuales,
        engine="openpyxl"
    )

    print(
        f"Registros Manuales: "
        f"{len(manuales):,}"
    )

    resultado = construir_vida(
        sap_vida,
        saa,
        manuales
    )

    print(
        f"Resultado VIDA: "
        f"{len(resultado):,}"
    )

    return resultado


def generar_gmm(
    ruta_saa,
    ruta_manuales
):

    print("\nCargando SAP GMM...")

    sap_gmm = pd.read_parquet(
        "gmm_para_reporte.parquet"
    )

    print(
        f"Registros SAP GMM: "
        f"{len(sap_gmm):,}"
    )

    print("\nCargando SAA...")

    saa = pd.read_excel(
        ruta_saa,
        engine="openpyxl"
    )

    print(
        f"Registros SAA: "
        f"{len(saa):,}"
    )

    print("\nCargando MANUALES...")

    manuales = pd.read_excel(
        ruta_manuales,
        engine="openpyxl"
    )

    print(
        f"Registros Manuales: "
        f"{len(manuales):,}"
    )

    resultado = construir_gmm(
        sap_gmm,
        saa,
        manuales
    )

    print(
        f"Resultado GMM: "
        f"{len(resultado):,}"
    )

    return resultado