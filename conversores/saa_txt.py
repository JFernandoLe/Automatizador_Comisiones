"""
Conversión de TXT SAA con delimitador '|'.

Acepta uno o varios archivos. La primera línea de cada archivo es el encabezado;
los encabezados deben coincidir para poder consolidar.
"""
import os
import tempfile
from pathlib import Path

from openpyxl import Workbook

MAX_FILAS_POR_HOJA = 1_048_576
ENCODING = "latin-1"
PREFIJO_HOJAS = "Datos"


def convertir_txt_saa(ruta_txt, encoding=ENCODING):
    return convertir_txts_saa([ruta_txt], encoding=encoding)


def convertir_txts_saa(rutas, encoding=ENCODING):
    archivos = [Path(ruta) for ruta in rutas if ruta]
    if not archivos:
        raise ValueError("Debe seleccionar al menos un archivo SAA TXT.")

    fd, ruta_salida = tempfile.mkstemp(prefix="saa_txt_", suffix=".xlsx")
    os.close(fd)

    wb = Workbook(write_only=True)
    encabezados = None
    hoja_actual = None
    numero_hoja = 1
    filas_en_hoja = 0
    total_registros = 0

    def crear_nueva_hoja():
        nonlocal hoja_actual, numero_hoja, filas_en_hoja
        hoja_actual = wb.create_sheet(f"{PREFIJO_HOJAS}_{numero_hoja}")
        numero_hoja += 1
        hoja_actual.append(encabezados)
        filas_en_hoja = 1

    for archivo in archivos:
        with open(archivo, "r", encoding=encoding) as f:
            encabezado_archivo = f.readline().strip()
            if not encabezado_archivo:
                raise ValueError(f"Archivo vacío: {archivo.name}")

            columnas = encabezado_archivo.split("|")
            if encabezados is None:
                encabezados = columnas
                crear_nueva_hoja()
            elif columnas != encabezados:
                raise ValueError(
                    f"El archivo {archivo.name} tiene encabezados distintos "
                    "al resto de los TXT SAA y no se puede consolidar."
                )

            for linea in f:
                linea = linea.strip()
                if not linea:
                    continue

                if filas_en_hoja >= MAX_FILAS_POR_HOJA:
                    crear_nueva_hoja()

                fila = linea.split("|")
                if len(fila) != len(encabezados):
                    print(
                        f"Cantidad de columnas incorrecta en {archivo.name}: "
                        f"{len(fila)} vs {len(encabezados)}"
                    )
                    continue

                hoja_actual.append(fila)
                filas_en_hoja += 1
                total_registros += 1
                if total_registros % 100000 == 0:
                    print(f"{total_registros:,} registros SAA TXT procesados...")

    if encabezados is None:
        raise ValueError("Ningún TXT SAA contenía encabezados.")

    wb.save(ruta_salida)
    print(
        f"SAA TXT convertido ({len(archivos)} archivo(s)): "
        f"{total_registros:,} registros -> {ruta_salida}"
    )
    return ruta_salida
