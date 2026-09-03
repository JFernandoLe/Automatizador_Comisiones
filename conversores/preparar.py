from pathlib import Path

import pandas as pd

from conversores.excel_original import (
    convertir_excel_original,
    convertir_exceles_originales,
    parece_ya_convertido,
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


def es_consolidado_listo(archivos):
    lista = como_lista(archivos)
    return len(lista) == 1 and parece_ya_convertido(lista[0])


def preparar_excel_sap(
    archivos,
    hojas_seleccionadas=None,
    prefijo="convertido_",
    actualizar_estado=None,
):
    """
    Originales → ConvExc (todas las hojas) → consolidado.
    Consolidado ya convertido → se deja igual para que el automatizador
    lo lea con pandas, sin +2 ni filtros otra vez.
    """
    lista = como_lista(archivos)
    if not lista:
        raise ValueError("Debe seleccionar al menos un archivo Excel.")

    if es_consolidado_listo(lista):
        print(f"Archivo ya convertido, sin reconvertir: {Path(lista[0]).name}")
        return lista[0]

    if len(lista) == 1:
        return convertir_excel_original(lista[0], prefijo=prefijo)

    return convertir_exceles_originales(
        lista,
        prefijo=prefijo,
        actualizar_estado=actualizar_estado,
    )


def leer_como_automatizador_original(archivos, ruta_preparada, hojas_seleccionadas):
    """
    Si el usuario trajo un consolidado, lee las hojas que marque.
    Si se acaba de consolidar, lee la primera hoja (pd.read_excel),
    como el automatizador leía Consolidado.xlsx.
    """
    lista = como_lista(archivos)
    if es_consolidado_listo(lista) and str(Path(ruta_preparada).resolve()) == str(
        Path(lista[0]).resolve()
    ):
        return leer_hojas_seleccionadas(ruta_preparada, hojas_seleccionadas)
    return pd.read_excel(ruta_preparada, engine="openpyxl")


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


def preparar_manuales(
    archivos,
    hojas_seleccionadas=None,
    actualizar_estado=None,
):
    lista = como_lista(archivos)
    if es_consolidado_listo(lista) and hojas_seleccionadas:
        df = leer_hojas_seleccionadas(lista[0], hojas_seleccionadas)
        return guardar_dataframe_temporal(df, prefijo="manuales_")

    return preparar_excel_sap(
        archivos,
        hojas_seleccionadas=hojas_seleccionadas,
        prefijo="manuales_",
        actualizar_estado=actualizar_estado,
    )
