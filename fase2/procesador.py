from pathlib import Path

import pandas as pd

from fase2.excel_salida import guardar_comision
from fase2.transformaciones import (
    construir_pagos_bonos,
    construir_tabla_gmm,
    construir_tabla_vida,
    crear_dist_map,
    crear_pfpm_map,
)
from servicios.excel import leer_hojas_seleccionadas


def _avisar(actualizar_estado, texto, progreso):
    if actualizar_estado:
        actualizar_estado(texto, progreso)


def _es_parquet(ruta):
    return Path(ruta).suffix.lower() == ".parquet"


def leer_reporte(ruta, hojas=None):
    if not ruta:
        raise ValueError("Debe seleccionar el reporte de la fase 1.")
    if _es_parquet(ruta):
        return pd.read_parquet(ruta)
    if not hojas:
        raise ValueError("Debe seleccionar al menos una hoja del reporte.")
    return leer_hojas_seleccionadas(ruta, hojas)


def _hoja_pfpm(hojas):
    if not hojas:
        raise ValueError("Debe seleccionar la hoja PFPM del archivo de catálogos.")
    for hoja in hojas:
        if str(hoja).strip().upper() == "PFPM":
            return hoja
    raise ValueError("Debe seleccionar la hoja PFPM del archivo de catálogos.")


def cargar_pfpm(ruta_catalogos, hojas_catalogos):
    if not ruta_catalogos:
        raise ValueError("Debe seleccionar el archivo de catálogos.")
    hoja = _hoja_pfpm(hojas_catalogos)
    df = leer_hojas_seleccionadas(ruta_catalogos, [hoja])
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]
    faltantes = [
        col for col in ("AGENTE_ORIGINAL", "AGENTE_REPORTERIA") if col not in df.columns
    ]
    if faltantes:
        raise ValueError(
            "La hoja PFPM no contiene las columnas obligatorias: "
            + ", ".join(faltantes)
        )
    return crear_pfpm_map(df)


def cargar_distribucion(ruta, hojas):
    if not ruta:
        raise ValueError("Debe seleccionar el archivo Distribución Comercial.")
    if not hojas:
        raise ValueError("Debe seleccionar al menos una hoja de Distribución Comercial.")
    df = leer_hojas_seleccionadas(ruta, hojas)
    return crear_dist_map(df)


def leer_bonos(ruta, hojas):
    if not ruta:
        raise ValueError("Debe seleccionar el archivo Bonos.")
    if not hojas:
        raise ValueError("Debe seleccionar al menos una hoja de Bonos.")
    return leer_hojas_seleccionadas(ruta, hojas, header=1)


def generar_comision_fase2(
    ruta_vida,
    hojas_vida,
    ruta_gmm,
    hojas_gmm,
    ruta_dist,
    hojas_dist,
    ruta_catalogos,
    hojas_catalogos,
    ruta_bonos,
    hojas_bonos,
    actualizar_estado=None,
    ruta_salida="Comision.xlsx",
):
    _avisar(actualizar_estado, "Leyendo Reporte VIDA Final...", 6)
    df_vida = leer_reporte(ruta_vida, hojas_vida)
    _avisar(actualizar_estado, "Leyendo Reporte GMM Final...", 16)
    df_gmm = leer_reporte(ruta_gmm, hojas_gmm)
    _avisar(actualizar_estado, "Leyendo Distribución Comercial...", 28)
    dist_map = cargar_distribucion(ruta_dist, hojas_dist)
    _avisar(actualizar_estado, "Leyendo catálogo PFPM...", 38)
    pfpm_map = cargar_pfpm(ruta_catalogos, hojas_catalogos)
    _avisar(actualizar_estado, "Leyendo archivo Bonos...", 48)
    df_bonos = leer_bonos(ruta_bonos, hojas_bonos)

    _avisar(actualizar_estado, "Construyendo tabla VIDA...", 60)
    tabla_vida = construir_tabla_vida(df_vida, dist_map, pfpm_map)
    _avisar(actualizar_estado, "Construyendo tabla GMM...", 72)
    tabla_gmm = construir_tabla_gmm(df_gmm, dist_map, pfpm_map)
    _avisar(actualizar_estado, "Procesando Pagos de Bonos...", 84)
    tabla_bonos = construir_pagos_bonos(df_bonos, dist_map, pfpm_map)

    _avisar(actualizar_estado, "Generando Comision.xlsx...", 93)
    ruta = guardar_comision(tabla_vida, tabla_gmm, tabla_bonos, ruta_salida)
    _avisar(actualizar_estado, "Fase 2 generada", 100)
    print(
        f"VIDA filas: {len(tabla_vida):,} | GMM filas: {len(tabla_gmm):,} | "
        f"Bonos filas: {len(tabla_bonos):,}"
    )
    print(f"Archivo generado: {ruta}")
    return ruta
