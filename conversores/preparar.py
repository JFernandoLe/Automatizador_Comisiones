from pathlib import Path

import pandas as pd

from conversores.excel_original import (
    convertir_excel_original,
    convertir_exceles_originales,
    hojas_listas_para_pandas,
)
from conversores.saa_txt import convertir_txts_saa
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
    return bool(lista) and all(es_txt(archivo) for archivo in lista)


def preparar_excel_sap(
    archivos,
    hojas_seleccionadas=None,
    prefijo="convertido_",
    actualizar_estado=None,
):
    """
    Un archivo + hojas marcadas:
      - si esas hojas ya tienen Fecha/Agente en fila 0 → no reconvertir
      - si no → ConvExc solo de esas hojas
    Varios archivos/carpeta:
      - ConvExc de todas las hojas de cada archivo (como el conversor)
    """
    lista = como_lista(archivos)
    if not lista:
        raise ValueError("Debe seleccionar al menos un archivo Excel.")

    if len(lista) == 1:
        if not hojas_seleccionadas:
            raise ValueError("Debe seleccionar al menos una hoja.")
        if hojas_listas_para_pandas(lista[0], hojas_seleccionadas):
            print(
                "Hojas ya convertidas, lectura pandas sin ConvExc: "
                + ", ".join(hojas_seleccionadas)
            )
            return lista[0]
        return convertir_excel_original(
            lista[0],
            prefijo=prefijo,
            hojas_seleccionadas=hojas_seleccionadas,
        )

    return convertir_exceles_originales(
        lista,
        prefijo=prefijo,
        actualizar_estado=actualizar_estado,
    )


def leer_como_automatizador_original(archivos, ruta_preparada, hojas_seleccionadas):
    """
    Un consolidado: solo las hojas que el usuario marcó (como el app original).
    Un original recién convertido o un lote: primera hoja del temp.
    """
    lista = como_lista(archivos)
    if (
        len(lista) == 1
        and hojas_seleccionadas
        and hojas_listas_para_pandas(lista[0], hojas_seleccionadas)
        and str(Path(ruta_preparada).resolve()) == str(Path(lista[0]).resolve())
    ):
        return leer_hojas_seleccionadas(ruta_preparada, hojas_seleccionadas)
    return pd.read_excel(ruta_preparada, engine="openpyxl")


def preparar_saa(archivos, hojas_seleccionadas):
    lista = como_lista(archivos)
    if not lista:
        raise ValueError("Debe seleccionar el archivo SAA.")

    txts = [archivo for archivo in lista if es_txt(archivo)]
    excels = [archivo for archivo in lista if not es_txt(archivo)]

    if txts and not excels:
        return convertir_txts_saa(txts)

    dataframes = []
    if txts:
        ruta_txt = convertir_txts_saa(txts)
        dataframes.append(
            leer_hojas_seleccionadas(ruta_txt, obtener_hojas(ruta_txt))
        )

    if not hojas_seleccionadas:
        raise ValueError("Debe seleccionar al menos una hoja de SAA.")

    for archivo in excels:
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
    if (
        len(lista) == 1
        and hojas_seleccionadas
        and hojas_listas_para_pandas(lista[0], hojas_seleccionadas)
    ):
        df = leer_hojas_seleccionadas(lista[0], hojas_seleccionadas)
        return guardar_dataframe_temporal(df, prefijo="manuales_")

    return preparar_excel_sap(
        archivos,
        hojas_seleccionadas=hojas_seleccionadas,
        prefijo="manuales_",
        actualizar_estado=actualizar_estado,
    )
