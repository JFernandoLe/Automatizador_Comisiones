from pathlib import Path

import pandas as pd

from conversores.excel_original import (
    convertir_excel_original,
    convertir_exceles_originales,
)
from conversores.saa_txt import convertir_txt_saa
from servicios.excel import (
    guardar_dataframe_temporal,
    leer_hojas_seleccionadas,
    obtener_hojas,
)


def es_txt(ruta):
    return Path(ruta).suffix.lower() == ".txt"


def como_lista(archivos):
    if not archivos:
        return []
    if isinstance(archivos, (list, tuple)):
        return [a for a in archivos if a]
    return [archivos]


def es_entrada_txt(archivos):
    lista = como_lista(archivos)
    return len(lista) == 1 and es_txt(lista[0])


def preparar_excel_sap(
    archivos,
    hojas_seleccionadas,
    prefijo="convertido_",
    actualizar_estado=None,
):
    lista = como_lista(archivos)
    if not lista:
        raise ValueError("Debe seleccionar al menos un archivo Excel.")

    if len(lista) == 1:
        ruta = convertir_excel_original(lista[0], hojas_seleccionadas, prefijo=prefijo)
        if not ruta:
            raise ValueError("Ninguna de las hojas seleccionadas existe en el archivo.")
        return ruta

    return convertir_exceles_originales(
        lista,
        hojas_seleccionadas,
        prefijo=prefijo,
        actualizar_estado=actualizar_estado,
    )


def preparar_saa(archivos, hojas_seleccionadas):
    lista = como_lista(archivos)
    if not lista:
        raise ValueError("Debe seleccionar el archivo SAA.")

    if es_entrada_txt(lista):
        return convertir_txt_saa(lista[0])

    if any(es_txt(archivo) for archivo in lista):
        raise ValueError(
            "SAA en TXT debe ser un solo archivo. "
            "No combine TXT con Excel ni seleccione varios TXT."
        )

    if not hojas_seleccionadas:
        raise ValueError("Debe seleccionar al menos una hoja de SAA.")

    dataframes = []
    for archivo in lista:
        hojas_archivo = [
            hoja for hoja in hojas_seleccionadas if hoja in obtener_hojas(archivo)
        ]
        if not hojas_archivo:
            print(f"SAA sin hojas seleccionadas: {Path(archivo).name}")
            continue
        dataframes.append(leer_hojas_seleccionadas(archivo, hojas_archivo))

    if not dataframes:
        raise ValueError("Ningún Excel SAA contenía las hojas seleccionadas.")

    return guardar_dataframe_temporal(
        pd.concat(dataframes, ignore_index=True),
        prefijo="saa_excel_",
    )


def preparar_manuales(archivos, hojas_seleccionadas, actualizar_estado=None):
    return preparar_excel_sap(
        archivos,
        hojas_seleccionadas,
        prefijo="manuales_",
        actualizar_estado=actualizar_estado,
    )
