from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from fase2.transformaciones import (
    CLASIF_APOYO,
    CLASIF_BONO,
    CLASIF_CONVENIO,
    clasificar_nueva_clasif,
    _figura_resumen,
    _ini_ren_resumen,
    _norm_nombre,
)

COLUMNAS_DATOS = [
    "Año",
    "Mes",
    "Conducto",
    "Promotor",
    "Clave",
    "Promotoria",
    "Concepto",
    "Ramo",
    "Ini_Ren",
    "Monto",
    "Fuente",
]

CONCEPTO_ANP = "ANP"
CONCEPTO_PRIMA = "PRIMA"
CONCEPTO_COMISION = "COMISION"
CONCEPTO_BONO = "BONO"
CONCEPTO_APOYO = "APOYO"
CONCEPTO_CONVENIO = "CONVENIO"

INI_PRIMER = "1ER AÑO"
INI_RENOV = "RENOVACION"
INI_OTROS = "OTROS CONCEPTOS"

MESES_NUM = {
    "ENERO": 1,
    "ENE": 1,
    "FEBRERO": 2,
    "FEB": 2,
    "MARZO": 3,
    "MAR": 3,
    "ABRIL": 4,
    "ABR": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JUN": 6,
    "JULIO": 7,
    "JUL": 7,
    "AGOSTO": 8,
    "AGO": 8,
    "SEPTIEMBRE": 9,
    "SEP": 9,
    "OCTUBRE": 10,
    "OCT": 10,
    "NOVIEMBRE": 11,
    "NOV": 11,
    "DICIEMBRE": 12,
    "DIC": 12,
}

MAPA_PRIMAS = (
    ("METPLUS", "GMM", INI_PRIMER),
    ("PRIMORDIAL1ER", "GMM", INI_PRIMER),
    ("PRIMORDIALRENOV", "GMM", INI_RENOV),
    ("GMMPRIMER", "GMM", INI_PRIMER),
    ("GMMRENOV", "GMM", INI_RENOV),
    ("VIDAPRIMER", "VIDA", INI_PRIMER),
    ("VIDARENOV", "VIDA", INI_RENOV),
)


def _avisar(actualizar_estado, texto, progreso=None):
    if actualizar_estado:
        actualizar_estado(texto, progreso)


def _a_entero(valor):
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        return None


def _mes_desfasado(anio, mes):
    if mes is None:
        return anio, mes
    mes = int(mes)
    if mes <= 1:
        return (int(anio) - 1 if anio is not None else None), 12
    return anio, mes - 1


def _mes_numero(valor):
    entero = _a_entero(valor)
    if entero is not None and 1 <= entero <= 12:
        return entero
    texto = _norm_nombre(valor).replace(".", "")
    return MESES_NUM.get(texto)


def _ramo_atomico(valor):
    entero = _a_entero(valor)
    if entero == 300:
        return "GMM"
    if entero == 101:
        return "VIDA"
    texto = _norm_nombre(valor)
    if "GMM" in texto:
        return "GMM"
    if "VIDA" in texto:
        return "VIDA"
    return None


def _es_valor_catalogo_ramo(valor):
    texto = _norm_nombre(valor).replace(" ", "").replace(".", "")
    if texto in {"101", "300"}:
        return False
    return "1VIDA" in texto or "2GMM" in texto or texto in {"VIDA", "GMM"}


def _ramo(valor):
    if isinstance(valor, pd.Series):
        catalogo = None
        original = None
        for item in valor.tolist():
            ramo = _ramo_atomico(item)
            if not ramo:
                continue
            if _es_valor_catalogo_ramo(item):
                catalogo = ramo
            else:
                original = ramo
        return catalogo or original
    return _ramo_atomico(valor)


def _es_columna_ramo(nombre):
    compacto = _norm_nombre(nombre).replace(".", "").replace(" ", "").replace("_", "")
    if compacto == "RAMOCLASIF":
        return True
    if compacto == "RAMO":
        return True
    return compacto.startswith("RAMO") and compacto[4:].isdigit()


def _indices_ramo(columnas):
    return [indice for indice, nombre in enumerate(columnas) if _es_columna_ramo(nombre)]


def _items_celda(valor):
    if isinstance(valor, pd.Series):
        return valor.tolist()
    return [valor]


def _clasificar_ramos(valores):
    catalogos = []
    originales = []
    for valor in valores:
        for item in _items_celda(valor):
            ramo = _ramo_atomico(item)
            if not ramo:
                continue
            if _es_valor_catalogo_ramo(item):
                catalogos.append(ramo)
            else:
                originales.append(ramo)
    catalogo = catalogos[-1] if catalogos else None
    original = originales[-1] if originales else None
    return catalogo, original


def _ramo_fila_bonos(fila, col_catalogo=None, columnas_ramo=None):
    """Columna X (2GMM/1Vida) gana sobre la G (101/300)."""
    if columnas_ramo:
        valores = [fila.iloc[indice] if indice < len(fila) else None for indice in columnas_ramo]
    else:
        valores = [fila.iloc[indice] for indice in _indices_ramo(fila.index)]
    catalogo, original = _clasificar_ramos(valores)
    return catalogo or original


def _parsear_fecha_bono(valor):
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, datetime):
        return valor
    if isinstance(valor, pd.Timestamp):
        if pd.isna(valor):
            return None
        return valor.to_pydatetime()
    numero = None
    try:
        if not isinstance(valor, str):
            numero = float(valor)
    except (TypeError, ValueError):
        numero = None
    if numero is not None and 20000 <= numero <= 80000:
        fecha = pd.to_datetime(numero, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(fecha):
            return fecha.to_pydatetime()
    fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.notna(fecha):
        anio = int(fecha.year)
        if 1990 <= anio <= 2100:
            return fecha.to_pydatetime()
    texto = str(valor).strip()
    for formato in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(texto[:10], formato)
        except (TypeError, ValueError):
            continue
    return None


def _hoja_pagos_bonos(ruta):
    try:
        hojas = _hojas(ruta)
    except Exception:
        return None
    for hoja in hojas:
        compacto = _norm_nombre(hoja).replace(" ", "").replace("_", "")
        if compacto == "PAGOSDEBONOS" or compacto.startswith("PAGOSDEBONOS"):
            return hoja
    return None


def _recortar_tabla_bonos(df):
    if df is None or df.empty:
        return df
    ultimo_dato = None
    for indice, col in enumerate(df.columns):
        compacto = _norm_nombre(col).replace(".", "").replace(" ", "").replace("_", "")
        if compacto == "NUEVACLASIF":
            return df.iloc[:, : indice + 1]
        if compacto in {
            "IMPORTE",
            "PROMOTORIA",
            "PRIMA",
            "FECHA",
            "MES",
            "PROMAGTE",
            "INIREN",
            "RAMO",
            "RAMO1",
            "RAMOCLASIF",
            "CPTO",
        } or (compacto.startswith("RAMO") and (len(compacto) == 4 or compacto[4:].isdigit())):
            ultimo_dato = indice
    if ultimo_dato is None:
        return df
    return df.iloc[:, : ultimo_dato + 1]


def _es_agente(valor):
    figura = _figura_resumen(valor)
    return figura in {"AGENTE", "AGENTE CVD"}


def _ini_tabla(valor):
    texto = _norm_nombre(valor)
    if "OTRO" in texto:
        return INI_OTROS
    if "RENOV" in texto or "CONSERV" in texto:
        return INI_RENOV
    if "1ER" in texto or "PRIMER" in texto:
        return INI_PRIMER
    return None


def _es_promotor(valor):
    return _figura_resumen(valor) == "PROMOTOR"


def _promotor(valor):
    entero = _a_entero(valor)
    if entero is not None:
        return entero
    texto = str(valor).strip()
    if not texto or texto.upper() in {"NAN", "NONE", "NAT", "N/A"}:
        return None
    return texto


def _completar_columnas(tabla):
    if tabla is None or tabla.empty:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    trabajo = tabla.copy()
    if "Promotoria" not in trabajo.columns:
        trabajo["Promotoria"] = trabajo.get("Promotor")
    if "Promotor" not in trabajo.columns:
        trabajo["Promotor"] = trabajo.get("Promotoria")
    for col in COLUMNAS_DATOS:
        if col not in trabajo.columns:
            trabajo[col] = None
    trabajo["Monto"] = pd.to_numeric(trabajo["Monto"], errors="coerce").fillna(0)
    trabajo["Año"] = trabajo["Año"].map(_a_entero)
    trabajo["Mes"] = trabajo["Mes"].map(_mes_numero)
    trabajo["Promotor"] = trabajo["Promotor"].map(_promotor)
    trabajo["Promotoria"] = trabajo["Promotoria"].map(_promotor)
    trabajo["Clave"] = trabajo["Clave"].map(_promotor)
    return trabajo[COLUMNAS_DATOS]


def _filas(registros, agrupar=True, exigir_promotor=True):
    if registros is None:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    if isinstance(registros, pd.DataFrame):
        if registros.empty:
            return pd.DataFrame(columns=COLUMNAS_DATOS)
        tabla = _completar_columnas(registros)
    else:
        if not registros:
            return pd.DataFrame(columns=COLUMNAS_DATOS)
        tabla = _completar_columnas(pd.DataFrame(registros))
    if agrupar:
        tabla = tabla[tabla["Monto"] != 0]
    tabla = tabla[tabla["Año"].notna() & tabla["Mes"].notna()]
    if exigir_promotor:
        tabla = tabla[tabla["Promotoria"].notna() | tabla["Promotor"].notna()]
    if tabla.empty:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    if not agrupar:
        return tabla.reset_index(drop=True)
    claves = [
        "Año",
        "Mes",
        "Conducto",
        "Promotor",
        "Clave",
        "Promotoria",
        "Concepto",
        "Ramo",
        "Ini_Ren",
        "Fuente",
    ]
    return tabla.groupby(claves, dropna=False, as_index=False)["Monto"].sum()


def _hojas(ruta):
    wb = load_workbook(ruta, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _tiene_hoja(ruta, nombre):
    objetivo = _norm_nombre(nombre)
    return any(_norm_nombre(hoja) == objetivo for hoja in _hojas(ruta))


def leer_datos_existente(ruta):
    try:
        hojas = _hojas(ruta)
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    for hoja in hojas:
        if _norm_nombre(hoja) == "DATOS":
            df = pd.read_excel(ruta, sheet_name=hoja)
            columnas = {_norm_nombre(col): col for col in df.columns}
            faltantes = [
                col
                for col in ("AÑO", "MES", "PROMOTOR", "CONCEPTO", "RAMO", "MONTO")
                if col not in columnas
            ]
            if faltantes:
                return pd.DataFrame(columns=COLUMNAS_DATOS)
            ini = columnas.get("INI_REN") or columnas.get("INI REN")
            trabajo = pd.DataFrame(
                {
                    "Año": df[columnas["AÑO"]].map(_a_entero),
                    "Mes": df[columnas["MES"]].map(_mes_numero),
                    "Conducto": df[columnas["CONDUCTO"]] if "CONDUCTO" in columnas else None,
                    "Promotor": df[columnas["PROMOTOR"]].map(_promotor),
                    "Clave": df[columnas["CLAVE"]].map(_promotor) if "CLAVE" in columnas else None,
                    "Promotoria": df[
                        columnas.get("PROMOTORIA", columnas["PROMOTOR"])
                    ].map(_promotor),
                    "Concepto": df[columnas["CONCEPTO"]].map(
                        lambda valor: str(valor).strip().upper()
                    ),
                    "Ramo": df[columnas["RAMO"]].map(_ramo),
                    "Ini_Ren": (
                        df[ini].map(_ini_tabla)
                        if ini
                        else pd.Series([None] * len(df))
                    ),
                    "Monto": pd.to_numeric(df[columnas["MONTO"]], errors="coerce"),
                }
            )
            return _filas(trabajo, agrupar=False, exigir_promotor=False)
    return pd.DataFrame(columns=COLUMNAS_DATOS)


def extraer_comision_base(ruta, anio):
    try:
        hojas = [
            hoja
            for hoja in _hojas(ruta)
            if _norm_nombre(hoja).startswith("COMISION BASE")
        ]
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    if not hojas:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    bloques = []
    for hoja in hojas:
        df = pd.read_excel(ruta, sheet_name=hoja)
        if df.empty:
            continue
        columnas = {_norm_nombre(col): col for col in df.columns}
        if "COMISION" not in columnas:
            continue
        nombre_hoja = _norm_nombre(hoja)
        anio_hoja = (int(anio) - 1) if "ANTERIOR" in nombre_hoja and anio else anio
        if "AÑO" in columnas:
            anios_col = df[columnas["AÑO"]].map(_a_entero)
        else:
            anios_col = anio_hoja
        col_tipo_com = columnas.get("TIPO COMISION")
        col_tipo = columnas.get("TIPO")
        col_promotor = columnas.get("PROMOTOR")
        col_promotoria = columnas.get("PROMOTORIA")
        col_clave = columnas.get("CLAVE")
        col_conducto = columnas.get("CONDUCTO")
        col_mes = columnas.get("MES")
        trabajo = pd.DataFrame(
            {
                "Año": anios_col,
                "Mes": df[col_mes].map(_mes_numero) if col_mes else None,
                "Conducto": df[col_conducto] if col_conducto else None,
                "Promotor": df[col_promotor].map(_promotor) if col_promotor else None,
                "Clave": df[col_clave].map(_promotor) if col_clave else None,
                "Promotoria": (
                    df[col_promotoria].map(_promotor) if col_promotoria else None
                ),
                "Concepto": CONCEPTO_COMISION,
                "Ramo": df[col_tipo].map(_ramo) if col_tipo else None,
                "Ini_Ren": df[col_tipo_com].map(_ini_tabla) if col_tipo_com else None,
                "Monto": df[columnas["COMISION"]],
                "Fuente": "Comision Base",
            }
        )
        if col_conducto:
            trabajo = trabajo[trabajo["Conducto"].map(_es_promotor)]
        bloques.append(trabajo)
    if not bloques:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    return _filas(pd.concat(bloques, ignore_index=True), agrupar=False, exigir_promotor=False)


def extraer_bonos(ruta, anio):
    hoja = _hoja_pagos_bonos(ruta)
    if not hoja:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    df = pd.read_excel(ruta, sheet_name=hoja)
    df = _recortar_tabla_bonos(df)
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    columnas = {_norm_nombre(col): col for col in df.columns}
    col_prom = columnas.get("PROMOTORIA")
    col_prima = columnas.get("PRIMA")
    col_importe = columnas.get("IMPORTE")
    if not col_importe:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    if not col_prom and not col_prima:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    col_ini = columnas.get("INI_REN")
    col_clasif = columnas.get("NUEVA CLASIF")
    col_desc = None
    for nombre, col in columnas.items():
        compacto = nombre.replace(".", "").replace(" ", "")
        if "DESC" in compacto and "CPTO" in compacto:
            col_desc = col
            break
        if nombre == "DESCCONCEPTO":
            col_desc = col
    col_fecha = columnas.get("FECHA")
    col_mes = columnas.get("MES")
    col_figura = columnas.get("PROM_AGTE") or columnas.get("PROMAGTE")
    idx_ramo = _indices_ramo(df.columns)
    if df.shape[1] > 23 and 23 not in idx_ramo:
        idx_ramo.append(23)
    idx_figura = None
    for indice, nombre in enumerate(df.columns):
        if _norm_nombre(nombre).replace(" ", "") in {"PROM_AGTE", "PROMAGTE"}:
            idx_figura = indice
    registros = []
    for pos in range(len(df)):
        fila = df.iloc[pos]
        valores_ramo = [df.iat[pos, indice] for indice in idx_ramo if indice < df.shape[1]]
        catalogo, original = _clasificar_ramos(valores_ramo)
        ramo = catalogo or original
        if not ramo:
            continue
        figura = df.iat[pos, idx_figura] if idx_figura is not None else (
            fila.get(col_figura) if col_figura else None
        )
        if _es_agente(figura) and not (catalogo == "GMM" and original == "VIDA"):
            continue
        ini_raw = fila.get(col_ini) if col_ini else None
        desc = fila.get(col_desc) if col_desc else ""
        clasif = fila.get(col_clasif) if col_clasif else None
        if not clasif or str(clasif).strip() in {"", "nan", "None"}:
            clasif = clasificar_nueva_clasif(ini_raw, desc)
        clasif_n = _norm_nombre(clasif)
        if clasif_n.startswith("APOYO"):
            concepto = CONCEPTO_APOYO
            ini = None
        elif clasif_n.startswith("CONVENIO"):
            concepto = CONCEPTO_CONVENIO
            ini = None
        else:
            concepto = CONCEPTO_BONO
            resumido = _ini_ren_resumen(ini_raw)
            if resumido == "Primer Año":
                ini = INI_PRIMER
            elif resumido == "Conservación":
                ini = INI_RENOV
            else:
                ini = INI_PRIMER
        anio_fila = None
        mes = None
        if col_fecha:
            fecha = _parsear_fecha_bono(fila.get(col_fecha))
            if fecha is not None:
                anio_fila = int(fecha.year)
                mes = int(fecha.month)
        if mes is None and col_mes:
            mes = _mes_numero(fila.get(col_mes))
        if anio_fila is None:
            anio_fila = anio
        anio_fila, mes = _mes_desfasado(anio_fila, mes)
        clave_prom = _promotor(fila.get(col_prom) if col_prom else None)
        if clave_prom is None and col_prima:
            clave_prom = _promotor(fila.get(col_prima))
        if clave_prom is None:
            continue
        registros.append(
            {
                "Año": anio_fila,
                "Mes": mes,
                "Conducto": "PROMOTOR",
                "Promotor": clave_prom,
                "Promotoria": clave_prom,
                "Concepto": concepto,
                "Ramo": ramo,
                "Ini_Ren": ini,
                "Monto": fila.get(col_importe),
                "Fuente": "Pagos_de_Bonos",
            }
        )
    return _filas(registros, exigir_promotor=False)


def extraer_anp(ruta):
    hojas = []
    try:
        hojas = _hojas(ruta)
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    nombre = None
    for hoja in hojas:
        if _norm_nombre(hoja) == "DETALLE":
            nombre = hoja
            break
    if nombre is None:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    df = pd.read_excel(ruta, sheet_name=nombre)
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    columnas = {_norm_nombre(col): col for col in df.columns}
    col_prom = columnas.get("CLAVE_PROMOTOR") or columnas.get("PROMOTOR")
    col_anio = columnas.get("ANIO_PAGO") or columnas.get("AÑO_PAGO") or columnas.get("AÑO")
    col_mes = columnas.get("MES_PAGO") or columnas.get("MES")
    col_ramo = columnas.get("RAMO_ANP") or columnas.get("RAMO")
    col_monto = columnas.get("PNA")
    if not all([col_prom, col_anio, col_mes, col_ramo, col_monto]):
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    registros = []
    for _, fila in df.iterrows():
        ramo = _ramo(fila.get(col_ramo))
        if not ramo:
            continue
        clave_prom = _promotor(fila.get(col_prom))
        registros.append(
            {
                "Año": _a_entero(fila.get(col_anio)),
                "Mes": _mes_numero(fila.get(col_mes)),
                "Promotor": clave_prom,
                "Promotoria": clave_prom,
                "Concepto": CONCEPTO_ANP,
                "Ramo": ramo,
                "Ini_Ren": None,
                "Monto": fila.get(col_monto),
                "Fuente": "ANP",
            }
        )
    return _filas(registros)


def _clave_titulo(titulo):
    texto = _norm_nombre(titulo)
    for sep in (" ", "-", "_", ".", "/"):
        texto = texto.replace(sep, "")
    return texto


def _mapear_titulo_prima(titulo):
    clave = _clave_titulo(titulo)
    for fragmento, ramo, ini in MAPA_PRIMAS:
        if fragmento in clave:
            return ramo, ini
    if "METPLUS" in clave:
        return "GMM", INI_PRIMER
    if "PRIMORDIAL" in clave:
        return "GMM", INI_RENOV if "RENOV" in clave else INI_PRIMER
    if "GMM" in clave:
        return "GMM", INI_RENOV if "RENOV" in clave else INI_PRIMER
    if "VIDA" in clave:
        return "VIDA", INI_RENOV if "RENOV" in clave else INI_PRIMER
    return None


def extraer_primas(ruta):
    if not Path(ruta).exists():
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    wb = load_workbook(ruta, data_only=True, read_only=True)
    if "Bases PP" not in wb.sheetnames:
        wb.close()
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    ws = wb["Bases PP"]
    filas = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(filas) < 3:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    encabezado = filas[1]
    registros = []
    for indice, valor in enumerate(encabezado):
        if _norm_nombre(valor) != "PROMOTOR":
            continue
        anio = _a_entero(filas[0][indice - 1] if indice else None)
        titulo = filas[0][indice] if indice < len(filas[0]) else None
        mapeo = _mapear_titulo_prima(titulo or "")
        if not mapeo or anio is None:
            continue
        ramo, ini = mapeo
        meses = {}
        col = indice + 2
        while col < len(encabezado):
            mes = _a_entero(encabezado[col])
            if mes is None or not 1 <= mes <= 12:
                break
            meses[col] = mes
            col += 1
        for fila in filas[2:]:
            if indice >= len(fila):
                continue
            promotor = _promotor(fila[indice])
            if promotor is None:
                continue
            for col_mes, mes in meses.items():
                if col_mes >= len(fila):
                    continue
                registros.append(
                    {
                        "Año": anio,
                        "Mes": mes,
                        "Promotor": promotor,
                        "Promotoria": promotor,
                        "Concepto": CONCEPTO_PRIMA,
                        "Ramo": ramo,
                        "Ini_Ren": ini,
                        "Monto": fila[col_mes],
                        "Fuente": str(titulo).strip() if titulo else "Bases PP",
                    }
                )
    return _filas(registros)


def extraer_paquete(ruta, anio, actualizar_estado=None):
    if not ruta or not Path(ruta).exists():
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    _avisar(actualizar_estado, f"Leyendo {Path(ruta).name}...", 20)
    existente = leer_datos_existente(ruta)
    partes = [existente] if not existente.empty else []
    partes.append(extraer_comision_base(ruta, anio))
    partes.append(extraer_bonos(ruta, anio))
    partes.append(extraer_anp(ruta))
    partes.append(extraer_primas(ruta))
    partes = [parte for parte in partes if parte is not None and not parte.empty]
    if not partes:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    return _completar_columnas(pd.concat(partes, ignore_index=True))


def consolidar_datos(paquetes, actualizar_estado=None):
    partes = []
    for anio, ruta in paquetes:
        if not ruta:
            continue
        partes.append(extraer_paquete(ruta, anio, actualizar_estado))
    partes = [parte for parte in partes if not parte.empty]
    if not partes:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    return _completar_columnas(pd.concat(partes, ignore_index=True))
