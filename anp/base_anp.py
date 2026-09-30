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
PRODUCTOS_GMM = {"gmm", "metplus", "primordial"}
PRODUCTOS_VIDA = {"metalife", "metalifec", "met4u", "vida tradicional"}
RAMOS_RESUMEN = ("GMM", "Vida")
RAIZ_PROYECTO = Path(__file__).resolve().parents[1]
RUTA_SALIDA_ANP = RAIZ_PROYECTO / "Base ANP.xlsx"
RUTA_LOG_ANP = RAIZ_PROYECTO / "LOG_ERRORES.xlsx"


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


def _compacto_columna(col):
    return re.sub(r"[^a-z0-9]+", " ", _norm_texto(col)).strip()


def encontrar_columna_pago(columnas):
    for col in columnas:
        compacto = _compacto_columna(col)
        for patron in PATRONES_PAGO:
            if patron in compacto or patron.replace(" ", "") in compacto.replace(" ", ""):
                return col
    return None


def encontrar_columna_por_patrones(columnas, patrones, exacto=False):
    for col in columnas:
        compacto = _compacto_columna(col)
        tokens = set(compacto.split())
        for patron in patrones:
            if exacto:
                if compacto == patron or tokens == set(patron.split()):
                    return col
            elif patron == compacto or patron in compacto:
                return col
    return None


def encontrar_columna_promotor(columnas):
    return encontrar_columna_por_patrones(
        columnas, ("clave promotor", "clavepromotor")
    ) or encontrar_columna_por_patrones(columnas, ("promotor",), exacto=True)


def encontrar_columna_pna(columnas):
    for col in columnas:
        compacto = _compacto_columna(col)
        if compacto == "pna" or "pna" in compacto.split():
            return col
    return None


def encontrar_columna_producto(columnas):
    return encontrar_columna_por_patrones(columnas, ("producto",), exacto=True)


def ramo_desde_producto(valor):
    texto = re.sub(r"\s+", " ", _norm_texto(valor))
    if texto in PRODUCTOS_GMM:
        return "GMM"
    if texto in PRODUCTOS_VIDA:
        return "Vida"
    return "Flexilife"


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


def _clave_promotor(valor):
    if _es_vacio(valor):
        return 0
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        texto = str(valor).strip()
        return texto if texto else 0


def _ordenar_promotor(valor):
    try:
        return (0, int(float(str(valor))))
    except (TypeError, ValueError):
        return (1, str(valor))


def _resolver_columna(df, encontrada, etiqueta, ruta, pedir_columna):
    if encontrada is not None:
        return encontrada
    if not pedir_columna:
        raise ValueError(f"No se localizó la columna {etiqueta}.")
    elegida = pedir_columna(
        Path(ruta).name, [str(col) for col in df.columns], etiqueta
    )
    if not elegida:
        raise ValueError(
            f"No se indicó la columna {etiqueta} para {Path(ruta).name}."
        )
    return elegida


def procesar_archivo_anp(ruta, anio, mes, hojas_preferidas, pedir_columna=None):
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

    if anio is None or mes is None:
        raise ValueError("No se indicó el periodo del archivo.")

    col_pago = _resolver_columna(
        df, encontrar_columna_pago(df.columns), "F. Pago", ruta, pedir_columna
    )
    col_promotor = _resolver_columna(
        df,
        encontrar_columna_promotor(df.columns),
        "Clave Promotor",
        ruta,
        pedir_columna,
    )
    col_pna = _resolver_columna(
        df, encontrar_columna_pna(df.columns), "PNA", ruta, pedir_columna
    )
    col_producto = _resolver_columna(
        df, encontrar_columna_producto(df.columns), "Producto", ruta, pedir_columna
    )

    df[col_pago] = pd.to_datetime(df[col_pago], errors="coerce")
    df[col_pago] = df[col_pago].apply(
        lambda valor: ajustar_fecha_pago(valor, anio, mes)
    )
    df["Clave_Promotor"] = df[col_promotor].map(_clave_promotor)
    df["PNA"] = pd.to_numeric(df[col_pna], errors="coerce").fillna(0)
    df["Ramo_ANP"] = df[col_producto].map(ramo_desde_producto)
    df["Mes_Pago"] = df[col_pago].dt.month
    df["Anio_Pago"] = df[col_pago].dt.year
    df["Archivo_Origen"] = Path(ruta).name
    df["Ruta_Origen"] = str(Path(ruta).resolve())
    df["Año_Proceso"] = anio
    df["Mes_Proceso"] = mes
    df["Periodo"] = f"{anio}-{mes:02d}"
    df["Fecha_Carga"] = datetime.now()
    return df


def construir_resumen_anp(df):
    trabajo = df.copy()
    trabajo = trabajo[trabajo["Ramo_ANP"].isin(RAMOS_RESUMEN)]
    trabajo = trabajo[trabajo["Anio_Pago"].notna() & trabajo["Mes_Pago"].notna()]
    if trabajo.empty:
        return {"anios_multiples": False, "meses": [], "filas": []}

    trabajo["Anio_Pago"] = trabajo["Anio_Pago"].astype(int)
    trabajo["Mes_Pago"] = trabajo["Mes_Pago"].astype(int)
    sumas = trabajo.groupby(
        ["Clave_Promotor", "Anio_Pago", "Mes_Pago", "Ramo_ANP"], dropna=False
    )["PNA"].sum()

    pares = sorted(
        {(int(anio), int(mes)) for anio, mes in zip(trabajo["Anio_Pago"], trabajo["Mes_Pago"])}
    )
    if pares:
        inicio, fin = pares[0], pares[-1]
        meses = []
        anio, mes = inicio
        while (anio, mes) <= fin:
            meses.append((anio, mes))
            if mes == 12:
                anio += 1
                mes = 1
            else:
                mes += 1
    else:
        meses = []
    anios = {anio for anio, _mes in meses}
    filas = []
    for promotor in sorted(trabajo["Clave_Promotor"].unique(), key=_ordenar_promotor):
        valores = {}
        total_general = 0.0
        for anio, mes in meses:
            gmm = 0.0
            vida = 0.0
            clave_gmm = (promotor, anio, mes, "GMM")
            clave_vida = (promotor, anio, mes, "Vida")
            if clave_gmm in sumas.index:
                gmm = float(sumas.loc[clave_gmm])
            if clave_vida in sumas.index:
                vida = float(sumas.loc[clave_vida])
            total = gmm + vida
            valores[(anio, mes)] = {"GMM": gmm, "Vida": vida, "TOTAL": total}
            total_general += total
        filas.append(
            {
                "promotor": promotor,
                "valores": valores,
                "total_general": total_general,
            }
        )
    return {
        "anios_multiples": len(anios) > 1,
        "meses": meses,
        "filas": filas,
    }


def _avisar(actualizar_estado, texto, progreso):
    if actualizar_estado:
        actualizar_estado(texto, progreso)


def generar_base_anp(
    archivos,
    periodos,
    hojas=None,
    actualizar_estado=None,
    pedir_columna=None,
    ruta_salida=None,
):
    from anp.excel_anp import guardar_base_anp

    lista = [str(Path(archivo).resolve()) for archivo in archivos if archivo]
    if not lista:
        raise ValueError("Debe seleccionar al menos un Excel de Base ANP.")
    ruta_salida = Path(ruta_salida) if ruta_salida else RUTA_SALIDA_ANP
    log_errores = RUTA_LOG_ANP

    errores = []
    bloques = []
    total = len(lista)
    for indice, archivo in enumerate(lista, start=1):
        progreso = 5 + int(80 * indice / total)
        _avisar(
            actualizar_estado,
            f"Procesando {Path(archivo).name} ({indice} de {total})...",
            progreso,
        )
        try:
            anio, mes = periodos[archivo]
            df = procesar_archivo_anp(
                archivo, anio, mes, hojas, pedir_columna=pedir_columna
            )
        except Exception as error:
            errores.append({"Archivo": archivo, "Error": str(error)})
            continue
        bloques.append(df)

    if not bloques:
        if errores:
            pd.DataFrame(errores).to_excel(log_errores, index=False)
        raise ValueError(
            "No se pudo generar la Base ANP. Revise LOG_ERRORES.xlsx."
        )

    _avisar(actualizar_estado, "Armando tabla dinámica...", 90)
    detalle = pd.concat(bloques, ignore_index=True, sort=False)
    resumen = construir_resumen_anp(detalle)
    guardar_base_anp(detalle, resumen, ruta_salida)
    if errores:
        pd.DataFrame(errores).to_excel(log_errores, index=False)
    elif log_errores.exists():
        log_errores.unlink()

    _avisar(actualizar_estado, "Base ANP generada", 100)
    return {
        "procesados": len(bloques),
        "errores": len(errores),
        "acumulado": str(ruta_salida),
        "log": str(log_errores) if errores else None,
    }
