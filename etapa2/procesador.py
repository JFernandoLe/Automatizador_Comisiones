import pandas as pd

from etapa2.transformaciones import (
    construir_vida,
    construir_gmm
)
from servicios.excel import leer_hojas_seleccionadas, obtener_hojas


def _cargar_excel_preparado(ruta):
    return leer_hojas_seleccionadas(ruta, obtener_hojas(ruta))


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

    saa = _cargar_excel_preparado(ruta_saa)

    print(
        f"Registros SAA: "
        f"{len(saa):,}"
    )

    print("\nCargando MANUALES...")

    manuales = _cargar_excel_preparado(ruta_manuales)

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

    saa = _cargar_excel_preparado(ruta_saa)

    print(
        f"Registros SAA: "
        f"{len(saa):,}"
    )

    print("\nCargando MANUALES...")

    manuales = _cargar_excel_preparado(ruta_manuales)

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