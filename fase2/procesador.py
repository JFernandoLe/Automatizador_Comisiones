from pathlib import Path

import pandas as pd

from fase2.excel_salida import guardar_comision
from fase2.transformaciones import (
    COLUMNAS_COMISION_BASE,
    avisos_cpto_sin_catalogo,
    construir_comision_base,
    construir_pagos_bonos,
    construir_tabla_gmm,
    construir_tabla_vida,
    crear_clasif_map,
    crear_dist_map,
    crear_pfpm_map,
    _norm_nombre,
)
from servicios.excel import leer_hojas_seleccionadas, obtener_hojas


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


def _mapear_columnas_comision_base(df):
    columnas = {_norm_nombre(col): col for col in df.columns}
    mapeo = {}
    for destino in COLUMNAS_COMISION_BASE:
        origen = columnas.get(_norm_nombre(destino))
        if origen is None:
            return None
        mapeo[destino] = origen
    return df.rename(columns={origen: destino for destino, origen in mapeo.items()})[
        COLUMNAS_COMISION_BASE
    ]


def leer_comision_base_externa(ruta, hojas=None, etiqueta="Base de comisiones FY"):
    if not ruta:
        return None
    try:
        disponibles = obtener_hojas(ruta)
    except Exception as error:
        raise ValueError(
            f"No se pudo leer la {etiqueta}: {error}"
        ) from error
    if not disponibles:
        raise ValueError(f"La {etiqueta} no tiene hojas.")
    if hojas:
        candidatas = [hoja for hoja in hojas if hoja in disponibles]
    else:
        candidatas = [
            hoja
            for hoja in disponibles
            if _norm_nombre(hoja).startswith("COMISION BASE")
            and "ANTERIOR" not in _norm_nombre(hoja)
        ]
        if not candidatas:
            candidatas = list(disponibles)
    bloques = []
    for hoja in candidatas:
        df = pd.read_excel(ruta, sheet_name=hoja)
        if df is None or df.empty:
            continue
        tabla = _mapear_columnas_comision_base(df)
        if tabla is None or tabla.empty:
            continue
        bloques.append(tabla)
    if not bloques:
        raise ValueError(
            f"La {etiqueta} no tiene la hoja "
            "Comision Base (Conducto, Mes, TIPO COMISIÓN, Tipo, PROMOTOR, "
            "CLAVE, COMISION, PROMOTORIA)."
        )
    return pd.concat(bloques, ignore_index=True)


def cargar_clasificaciones(ruta, hojas):
    if not ruta:
        raise ValueError("Debe seleccionar el Catálogo de clasificaciones.")
    if not hojas:
        raise ValueError(
            "Debe seleccionar las hojas VIDA y GMM del Catálogo de clasificaciones."
        )
    df = leer_hojas_seleccionadas(ruta, hojas)
    return crear_clasif_map(df)


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
    ruta_clasif,
    hojas_clasif,
    combinaciones=None,
    actualizar_estado=None,
    ruta_salida="Comision.xlsx",
    ruta_bonos_anterior=None,
    hojas_bonos_anterior=None,
    ruta_comision_anterior=None,
    hojas_comision_anterior=None,
    ruta_bonos_dos_anios=None,
    hojas_bonos_dos_anios=None,
    ruta_comision_dos_anios=None,
    hojas_comision_dos_anios=None,
):
    _avisar(actualizar_estado, "Leyendo Reporte VIDA Final...", 6)
    df_vida = leer_reporte(ruta_vida, hojas_vida)
    _avisar(actualizar_estado, "Leyendo Reporte GMM Final...", 14)
    df_gmm = leer_reporte(ruta_gmm, hojas_gmm)
    _avisar(actualizar_estado, "Leyendo Distribución Comercial...", 24)
    dist_map = cargar_distribucion(ruta_dist, hojas_dist)
    _avisar(actualizar_estado, "Leyendo catálogo PFPM...", 32)
    pfpm_map = cargar_pfpm(ruta_catalogos, hojas_catalogos)
    _avisar(actualizar_estado, "Leyendo archivo Bonos...", 42)
    df_bonos = leer_bonos(ruta_bonos, hojas_bonos)
    df_bonos_anterior = None
    if ruta_bonos_anterior:
        _avisar(actualizar_estado, "Leyendo Bonos FY del año anterior...", 47)
        df_bonos_anterior = leer_bonos(ruta_bonos_anterior, hojas_bonos_anterior)
    df_bonos_dos = None
    if ruta_bonos_dos_anios:
        _avisar(actualizar_estado, "Leyendo Bonos FY de hace dos años...", 49)
        df_bonos_dos = leer_bonos(ruta_bonos_dos_anios, hojas_bonos_dos_anios)
    _avisar(actualizar_estado, "Leyendo Catálogo de clasificaciones...", 52)
    clasif_map = cargar_clasificaciones(ruta_clasif, hojas_clasif)

    _avisar(actualizar_estado, "Construyendo tabla VIDA...", 62)
    tabla_vida = construir_tabla_vida(df_vida, dist_map, pfpm_map)
    _avisar(actualizar_estado, "Construyendo tabla GMM...", 72)
    tabla_gmm = construir_tabla_gmm(df_gmm, dist_map, pfpm_map)
    _avisar(actualizar_estado, "Construyendo Comisión Base...", 78)
    tabla_base = construir_comision_base(df_vida, df_gmm, dist_map, pfpm_map)
    tabla_base_anterior = None
    if ruta_comision_anterior:
        _avisar(actualizar_estado, "Leyendo Comision Base FY del año anterior...", 81)
        tabla_base_anterior = leer_comision_base_externa(
            ruta_comision_anterior,
            hojas_comision_anterior,
            etiqueta="Base de comisiones FY del año anterior",
        )
    tabla_base_dos = None
    if ruta_comision_dos_anios:
        _avisar(actualizar_estado, "Leyendo Comision Base FY de hace dos años...", 82)
        tabla_base_dos = leer_comision_base_externa(
            ruta_comision_dos_anios,
            hojas_comision_dos_anios,
            etiqueta="Base de comisiones FY de hace dos años",
        )
    _avisar(actualizar_estado, "Procesando Pagos de Bonos...", 84)
    tabla_bonos = construir_pagos_bonos(df_bonos, dist_map, pfpm_map, clasif_map)
    avisos = avisos_cpto_sin_catalogo(tabla_bonos)
    if df_bonos_anterior is not None and not df_bonos_anterior.empty:
        tabla_bonos_ant = construir_pagos_bonos(
            df_bonos_anterior, dist_map, pfpm_map, clasif_map
        )
        avisos.extend(avisos_cpto_sin_catalogo(tabla_bonos_ant))
        tabla_bonos = pd.concat([tabla_bonos, tabla_bonos_ant], ignore_index=True)
    if df_bonos_dos is not None and not df_bonos_dos.empty:
        tabla_bonos_dos = construir_pagos_bonos(
            df_bonos_dos, dist_map, pfpm_map, clasif_map
        )
        avisos.extend(avisos_cpto_sin_catalogo(tabla_bonos_dos))
        tabla_bonos = pd.concat([tabla_bonos, tabla_bonos_dos], ignore_index=True)

    _avisar(actualizar_estado, "Generando Comision.xlsx...", 93)
    ruta = guardar_comision(
        tabla_vida,
        tabla_gmm,
        tabla_bonos,
        ruta_salida,
        combinaciones=combinaciones,
        df_base=tabla_base,
        df_base_anterior=tabla_base_anterior,
        df_base_dos_anios=tabla_base_dos,
    )
    _avisar(actualizar_estado, "Fase 2 generada", 100)
    print(
        f"VIDA filas: {len(tabla_vida):,} | GMM filas: {len(tabla_gmm):,} | "
        f"Bonos filas: {len(tabla_bonos):,}"
    )
    print(f"Archivo generado: {ruta}")
    return ruta, avisos
