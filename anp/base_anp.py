import calendar
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

EXTENSIONES_ANP = {".xlsx", ".xls", ".xlsm", ".xlsb"}
HOJA_DETALLE = "detalle pagado"
FILA_ENCABEZADO_RESPALDO = 8
MAX_FILAS_ENCABEZADO = 25

MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

MESES_INV = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

PATRONES_PAGO = (
    "f. pago",
    "f pago",
    "fecha pago",
    "fecha de pago",
    "fpago",
)


def _norm_texto(valor):
    return (
        str(valor)
        .strip()
        .lower()
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
    )


def _es_vacio(valor):
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    return texto == "" or texto.upper() in {"NAN", "NONE", "NAT"}


def detectar_periodo_archivo(ruta):
    nombre = Path(ruta).name.lower()
    mes_anio = re.match(r"^(\d{1,2})[_-]?(20\d{2})", nombre)
    if mes_anio:
        mes = int(mes_anio.group(1))
        anio = int(mes_anio.group(2))
        if 1 <= mes <= 12:
            return anio, mes

    anio_mes = re.search(r"(20\d{2})[\s_-]?(0?[1-9]|1[0-2])", nombre)
    if anio_mes:
        return int(anio_mes.group(1)), int(anio_mes.group(2))

    anio_match = re.search(r"(20\d{2})", nombre)
    anio = int(anio_match.group(1)) if anio_match else None
    for texto_mes, numero_mes in MESES.items():
        if texto_mes in nombre and anio is not None:
            return anio, numero_mes
    return None, None


def ajustar_fecha_pago(fecha, anio, mes):
    if pd.isna(fecha):
        return fecha
    try:
        fecha = pd.to_datetime(fecha)
        inicio_mes = datetime(anio, mes, 1)
        ultimo = calendar.monthrange(anio, mes)[1]
        fin_mes = datetime(anio, mes, ultimo)
        if fecha.year < anio:
            return inicio_mes
        if fecha.year > anio:
            return fin_mes
        if fecha.month < mes:
            return inicio_mes
        if fecha.month > mes:
            return fin_mes
        return fecha
    except (TypeError, ValueError, OverflowError):
        return fecha


def encontrar_columna_pago(columnas):
    for col in columnas:
        nombre = _norm_texto(col)
        compacto = re.sub(r"[^a-z0-9]+", " ", nombre).strip()
        for patron in PATRONES_PAGO:
            if patron in compacto or patron.replace(" ", "") in compacto.replace(" ", ""):
                return col
    return None


def _fila_tiene_pago(valores):
    return encontrar_columna_pago(valores) is not None


def detectar_fila_encabezado_anp(df):
    limite = min(len(df), MAX_FILAS_ENCABEZADO)
    for indice in range(limite):
        if _fila_tiene_pago(df.iloc[indice].tolist()):
            return indice
    if len(df) > FILA_ENCABEZADO_RESPALDO:
        return FILA_ENCABEZADO_RESPALDO
    return 0


def _encabezados_unicos(nombres):
    vistos = {}
    unicos = []
    for nombre in nombres:
        if nombre in vistos:
            vistos[nombre] += 1
            unicos.append(f"{nombre}_{vistos[nombre]}")
        else:
            vistos[nombre] = 0
            unicos.append(nombre)
    return unicos


def aplicar_encabezado_anp(df):
    if df is None or df.empty:
        return df
    fila = detectar_fila_encabezado_anp(df)
    encabezados = []
    for indice, valor in enumerate(df.iloc[fila].tolist()):
        if _es_vacio(valor):
            encabezados.append(f"Unnamed_{indice}")
        else:
            encabezados.append(str(valor).strip())
    datos = df.iloc[fila + 1 :].copy()
    datos.columns = _encabezados_unicos(encabezados)
    return datos.dropna(how="all").reset_index(drop=True)


def limpiar_filas_resumen(df):
    if df is None or df.empty:
        return df
    mask = df.astype(str).apply(lambda col: col.str.strip().str.upper())
    subtotal = mask.eq("SUBTOTAL").any(axis=1)
    if not subtotal.any():
        return df
    idx = subtotal.idxmax()
    if idx == 0:
        return df.iloc[0:0]
    return df.loc[: idx - 1]


def listar_excel_anp(origenes, incluir_subcarpetas=True):
    archivos = []
    vistos = set()
    for origen in origenes:
        ruta = Path(origen)
        if ruta.is_file():
            if ruta.suffix.lower() in EXTENSIONES_ANP and not ruta.name.startswith("~$"):
                resuelta = str(ruta.resolve())
                if resuelta not in vistos:
                    vistos.add(resuelta)
                    archivos.append(resuelta)
            continue
        if not ruta.is_dir():
            continue
        iterator = ruta.rglob("*") if incluir_subcarpetas else ruta.iterdir()
        for archivo in iterator:
            if not archivo.is_file():
                continue
            if archivo.suffix.lower() not in EXTENSIONES_ANP:
                continue
            if archivo.name.startswith("~$"):
                continue
            resuelta = str(archivo.resolve())
            if resuelta not in vistos:
                vistos.add(resuelta)
                archivos.append(resuelta)
    return sorted(archivos)


def _hojas_archivo(ruta, hojas_preferidas):
    from servicios.excel import obtener_hojas

    disponibles = obtener_hojas(ruta)
    if not disponibles:
        return []
    detalle = [
        hoja
        for hoja in disponibles
        if hoja.strip().lower() == HOJA_DETALLE
    ]
    if detalle:
        return detalle
    if hojas_preferidas:
        return [hoja for hoja in hojas_preferidas if hoja in disponibles]
    return []


def _leer_hojas_anp(ruta, hojas):
    from servicios.excel import leer_hoja_sin_encabezado

    dataframes = []
    for hoja in hojas:
        crudo = leer_hoja_sin_encabezado(ruta, hoja)
        datos = aplicar_encabezado_anp(crudo)
        if datos is not None and not datos.empty:
            dataframes.append(datos)
    if not dataframes:
        return None
    return pd.concat(dataframes, ignore_index=True)


def procesar_archivo_anp(ruta, anio_respaldo, mes_respaldo, hojas_preferidas):
    hojas = _hojas_archivo(ruta, hojas_preferidas)
    if not hojas:
        raise ValueError(
            "No se encontró la hoja 'Detalle pagado' ni hojas marcadas para este archivo."
        )
    df = _leer_hojas_anp(ruta, hojas)
    if df is None or df.empty:
        raise ValueError("No se encontraron datos en la hoja de detalle.")

    df = limpiar_filas_resumen(df)
    df = df.dropna(how="all").reset_index(drop=True)
    if df.empty:
        raise ValueError("El detalle quedó vacío después de quitar resumenes.")

    anio, mes = detectar_periodo_archivo(ruta)
    if anio is None or mes is None:
        anio, mes = anio_respaldo, mes_respaldo
    if anio is None or mes is None:
        raise ValueError("No se detectó el periodo en el nombre del archivo.")

    col_pago = encontrar_columna_pago(df.columns)
    if col_pago is None:
        raise ValueError("No se localizó la columna F. Pago / Fecha de pago.")

    df[col_pago] = pd.to_datetime(df[col_pago], errors="coerce")
    df[col_pago] = df[col_pago].apply(
        lambda valor: ajustar_fecha_pago(valor, anio, mes)
    )
    df["Archivo_Origen"] = Path(ruta).name
    df["Ruta_Origen"] = str(Path(ruta).resolve())
    df["Año_Proceso"] = anio
    df["Mes_Proceso"] = mes
    df["Periodo"] = f"{anio}-{mes:02d}"
    df["Fecha_Carga"] = datetime.now()
    return df


def _avisar(actualizar_estado, texto, progreso):
    if actualizar_estado:
        actualizar_estado(texto, progreso)


def generar_base_anp(
    archivos,
    carpeta_salida,
    anio_respaldo=None,
    mes_respaldo=None,
    hojas=None,
    actualizar_estado=None,
):
    lista = [str(Path(archivo).resolve()) for archivo in archivos if archivo]
    if not lista:
        raise ValueError("Debe seleccionar al menos un Excel de Base ANP.")
    if not carpeta_salida:
        raise ValueError("Debe seleccionar la carpeta de salida.")

    salida = Path(carpeta_salida)
    salida.mkdir(parents=True, exist_ok=True)
    carpeta_csv = salida / "CSV_Individuales"
    carpeta_csv.mkdir(exist_ok=True)
    acumulado_csv = salida / "Bases ANP.csv"
    log_errores = salida / "LOG_ERRORES.xlsx"
    if acumulado_csv.exists():
        acumulado_csv.unlink()

    errores = []
    bloques = []
    total = len(lista)
    for indice, archivo in enumerate(lista, start=1):
        progreso = 5 + int(85 * indice / total)
        _avisar(
            actualizar_estado,
            f"Procesando {Path(archivo).name} ({indice} de {total})...",
            progreso,
        )
        try:
            df = procesar_archivo_anp(
                archivo, anio_respaldo, mes_respaldo, hojas
            )
        except Exception as error:
            errores.append({"Archivo": archivo, "Error": str(error)})
            continue
        nombre = Path(archivo).stem
        df.to_csv(
            carpeta_csv / f"{nombre}.csv",
            index=False,
            encoding="utf-8-sig",
        )
        bloques.append(df)

    if bloques:
        pd.concat(bloques, ignore_index=True, sort=False).to_csv(
            acumulado_csv,
            index=False,
            encoding="utf-8-sig",
        )
    if errores:
        pd.DataFrame(errores).to_excel(log_errores, index=False)
    elif log_errores.exists():
        log_errores.unlink()

    _avisar(actualizar_estado, "Base ANP generada", 100)
    if not bloques:
        raise ValueError(
            "No se pudo generar la Base ANP. Revise LOG_ERRORES.xlsx."
        )
    return {
        "procesados": len(bloques),
        "errores": len(errores),
        "acumulado": str(acumulado_csv),
        "csv": str(carpeta_csv),
        "log": str(log_errores) if errores else None,
    }
