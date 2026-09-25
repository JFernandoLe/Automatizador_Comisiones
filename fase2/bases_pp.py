import re
from pathlib import Path

import pandas as pd

VALOR_NA = "N/A"
PATRON_MES_ANIO = re.compile(r"^(\d{1,2})[_-]?(\d{4})")

TIPOS_BASES_PP = (
    {
        "id": "vida_pa",
        "titulo": "Vida Primer Año",
        "titulo_corto": "Vida Primer Año",
    },
    {
        "id": "vida_ren",
        "titulo": "Vida Renovacion",
        "titulo_corto": "Vida Renovacion",
    },
    {
        "id": "gmm_pa",
        "titulo": "GMM Primer Año",
        "titulo_corto": "GMM Primer Año",
    },
    {
        "id": "gmm_ren",
        "titulo": "GMM Renovacion",
        "titulo_corto": "GMM Renovacion",
    },
    {
        "id": "prim_pa",
        "titulo": "Primordial 1er Año",
        "titulo_corto": "Primordial 1er Año",
    },
    {
        "id": "prim_ren",
        "titulo": "Primordial Renovacion",
        "titulo_corto": "Primordial Renovacion",
    },
    {
        "id": "metplus",
        "titulo": "Metplus",
        "titulo_corto": "Metplus",
    },
)
ORDEN_TIPOS_PP = tuple(item["id"] for item in TIPOS_BASES_PP)
TITULOS_PP = {item["id"]: item["titulo"] for item in TIPOS_BASES_PP}

COL_PROMOTOR_RES = 2
COL_AGENTE_RES = 5
COL_PRIMA_RES = 17
COL_MES_RES = 19


def _norm_texto(valor):
    return (
        str(valor)
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ñ", "N")
    )


def _clave_archivo(ruta):
    stem = _norm_texto(Path(ruta).stem)
    return re.sub(r"[^A-Z0-9]", "", stem)


def extraer_mes_anio(ruta):
    nombre = Path(ruta).name
    coincide = PATRON_MES_ANIO.match(nombre)
    if not coincide:
        clave = _clave_archivo(ruta)
        coincide = re.match(r"^(\d{1,2})(\d{4})", clave)
    if not coincide:
        raise ValueError(
            f"No se pudo obtener mes y año del archivo: {Path(ruta).name}"
        )
    mes = int(coincide.group(1))
    anio = int(coincide.group(2))
    if mes < 1 or mes > 12:
        raise ValueError(f"Mes inválido en el archivo: {Path(ruta).name}")
    return mes, anio


def clasificar_excel_pp(ruta):
    clave = _clave_archivo(ruta)
    clave = re.sub(r"^\d{1,2}\d{4}", "", clave)
    if "METPLUS" in clave:
        return "metplus"
    if "PRIMORDIAL" in clave and "RENOV" in clave:
        return "prim_ren"
    if "PRIMORDIAL" in clave:
        return "prim_pa"
    if "GMM" in clave and "RENOV" in clave:
        return "gmm_ren"
    if "GMM" in clave and ("PPONDERADA" in clave or "PONDERADA" in clave):
        return "gmm_pa"
    if "VIDA" in clave and "RENOV" in clave:
        return "vida_ren"
    if "VIDA" in clave and "PCA" in clave:
        return "vida_pa"
    if "VIDA" in clave:
        return "vida_pa"
    if "GMM" in clave:
        return "gmm_pa"
    return None


def listar_excel_primas(carpeta):
    from servicios.excel import listar_excel_en_carpeta

    if not carpeta:
        raise ValueError("Debe seleccionar la carpeta Bases de Primas del mes.")
    archivos = listar_excel_en_carpeta(carpeta, incluir_xlsb=True)
    if not archivos:
        raise ValueError(
            "La carpeta no contiene archivos Excel (.xlsx, .xls o .xlsb) en el primer nivel."
        )
    return archivos


def agrupar_excel_pp(archivos):
    grupos = {tipo: [] for tipo in ORDEN_TIPOS_PP}
    sin_tipo = []
    for archivo in archivos:
        tipo = clasificar_excel_pp(archivo)
        if tipo is None:
            sin_tipo.append(archivo)
            continue
        grupos[tipo].append(archivo)
    return grupos, sin_tipo


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


def _con_na(valor):
    return VALOR_NA if _es_vacio(valor) else valor


def _a_entero(valor):
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        return None


def _clave_tabla(valor):
    if _es_vacio(valor):
        return VALOR_NA
    entero = _a_entero(valor)
    if entero is not None:
        return entero
    texto = str(valor).strip()
    return texto if texto else VALOR_NA


def _ordenar_clave(valor):
    entero = _a_entero(valor)
    if entero is not None:
        return (0, entero)
    return (1, str(valor))


def _norm_columna(col):
    return re.sub(r"[^A-Z0-9]+", " ", _norm_texto(col)).strip()


def _buscar_columna_pp(df, nombres, indice_respaldo, etiqueta):
    objetivos = {_norm_columna(nombre) for nombre in nombres}
    exactas = []
    parciales = []
    for col in df.columns:
        actual = _norm_columna(col)
        if actual in objetivos:
            exactas.append(col)
        elif any(
            objetivo and objetivo in actual
            for objetivo in objetivos
            if len(objetivo) >= 4
        ):
            parciales.append(col)
    if len(exactas) == 1:
        return exactas[0]
    if len(exactas) > 1:
        return exactas[0]
    if len(parciales) == 1:
        return parciales[0]
    if df.shape[1] > indice_respaldo:
        return df.columns[indice_respaldo]
    raise ValueError(f"No se encontró la columna {etiqueta} en el Excel de primas.")


def _columnas_pp(df):
    return {
        "promotor": _buscar_columna_pp(
            df, ("Promotor",), COL_PROMOTOR_RES, "Promotor"
        ),
        "agente": _buscar_columna_pp(df, ("Agente",), COL_AGENTE_RES, "Agente"),
        "prima": _buscar_columna_pp(
            df,
            ("Prima Pagada", "Prima_Pagada", "PPagada"),
            COL_PRIMA_RES,
            "Prima Pagada",
        ),
        "mes": _buscar_columna_pp(df, ("Mes",), COL_MES_RES, "Mes"),
    }


def construir_tabla_pp(df, mes_hasta, anio, titulo):
    columnas = _columnas_pp(df)
    trabajo = pd.DataFrame(
        {
            "Promotor": df[columnas["promotor"]].map(_clave_tabla),
            "Agente": df[columnas["agente"]].map(_clave_tabla),
            "Mes": df[columnas["mes"]].map(_a_entero),
            "Prima": pd.to_numeric(df[columnas["prima"]], errors="coerce").fillna(0),
        }
    )
    trabajo = trabajo[
        trabajo["Mes"].between(1, mes_hasta) & (trabajo["Prima"] != 0)
    ]
    if trabajo.empty:
        return {
            "anio": anio,
            "titulo": titulo,
            "mes_hasta": mes_hasta,
            "filas": [],
        }

    sumas = trabajo.groupby(["Promotor", "Agente", "Mes"], dropna=False)[
        "Prima"
    ].sum()
    filas = []
    pares = sorted(
        {(promotor, agente) for promotor, agente, _mes in sumas.index},
        key=lambda par: (_ordenar_clave(par[0]), _ordenar_clave(par[1])),
    )
    for promotor, agente in pares:
        meses = {}
        total = 0.0
        for mes in range(1, mes_hasta + 1):
            clave = (promotor, agente, mes)
            monto = float(sumas.loc[clave]) if clave in sumas.index else 0.0
            meses[mes] = monto
            total += monto
        if total == 0:
            continue
        filas.append(
            {
                "promotor": promotor,
                "agente": agente,
                "meses": meses,
                "total": total,
            }
        )
    return {
        "anio": anio,
        "titulo": titulo,
        "mes_hasta": mes_hasta,
        "filas": filas,
    }


def _avisar(actualizar_estado, texto, progreso):
    if actualizar_estado:
        actualizar_estado(texto, progreso)


def generar_primas(entradas, ruta_salida="Primas.xlsx", actualizar_estado=None):
    from fase2.excel_primas import guardar_primas
    from servicios.excel import leer_hojas_seleccionadas

    if not entradas:
        raise ValueError("No hay archivos de primas seleccionados para generar.")
    tablas = []
    total = len(entradas)
    for indice, entrada in enumerate(entradas, start=1):
        progreso = 8 + int(80 * indice / total)
        _avisar(
            actualizar_estado,
            f"Leyendo {entrada['titulo']} {entrada['anio']}...",
            progreso,
        )
        df = leer_hojas_seleccionadas(entrada["archivo"], entrada["hojas"])
        tablas.append(
            construir_tabla_pp(
                df,
                entrada["mes_hasta"],
                entrada["anio"],
                entrada["titulo"],
            )
        )
    _avisar(actualizar_estado, "Generando Primas.xlsx...", 94)
    ruta = guardar_primas(tablas, ruta_salida)
    _avisar(actualizar_estado, "Bases PP generada", 100)
    print(f"Archivo generado: {ruta}")
    return ruta
