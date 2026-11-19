from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from fase2.transformaciones import (
    CLASIF_APOYO,
    CLASIF_BONO,
    CLASIF_CONVENIO,
    clasificar_nueva_clasif,
    _ini_ren_resumen,
    _norm_nombre,
)

COLUMNAS_DATOS = [
    "Año",
    "Mes",
    "Promotor",
    "Concepto",
    "Ramo",
    "Ini_Ren",
    "Monto",
]

CONCEPTO_ANP = "ANP"
CONCEPTO_PRIMA = "PRIMA"
CONCEPTO_COMISION = "COMISION"
CONCEPTO_BONO = "BONO"
CONCEPTO_APOYO = "APOYO"
CONCEPTO_CONVENIO = "CONVENIO"

INI_PRIMER = "1ER AÑO"
INI_RENOV = "RENOVACION"

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
    ("VIDA PRIMER AÑO", "VIDA", INI_PRIMER),
    ("VIDA RENOV", "VIDA", INI_RENOV),
    ("GMM PRIMER AÑO", "GMM", INI_PRIMER),
    ("GMM RENOV", "GMM", INI_RENOV),
    ("PRIMORDIAL 1ER", "GMM", INI_PRIMER),
    ("PRIMORDIAL RENOV", "GMM", INI_RENOV),
    ("METPLUS", "GMM", INI_PRIMER),
)


def _avisar(actualizar_estado, texto, progreso=None):
    if actualizar_estado:
        actualizar_estado(texto, progreso)


def _a_entero(valor):
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        return None


def _mes_numero(valor):
    entero = _a_entero(valor)
    if entero is not None and 1 <= entero <= 12:
        return entero
    texto = _norm_nombre(valor).replace(".", "")
    return MESES_NUM.get(texto)


def _ramo(valor):
    texto = _norm_nombre(valor)
    if "GMM" in texto or texto == "300":
        return "GMM"
    if "VIDA" in texto or texto == "101":
        return "VIDA"
    return None


def _ini_tabla(valor):
    texto = _norm_nombre(valor)
    if "OTRO" in texto:
        return None
    if "RENOV" in texto or "CONSERV" in texto:
        return INI_RENOV
    if "1ER" in texto or "PRIMER" in texto:
        return INI_PRIMER
    return None


def _promotor(valor):
    entero = _a_entero(valor)
    if entero is not None:
        return entero
    texto = str(valor).strip()
    if not texto or texto.upper() in {"NAN", "NONE", "NAT", "N/A"}:
        return None
    return texto


def _filas(registros):
    if not registros:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    tabla = pd.DataFrame(registros)
    tabla["Monto"] = pd.to_numeric(tabla["Monto"], errors="coerce").fillna(0)
    tabla = tabla[tabla["Monto"] != 0]
    tabla = tabla[tabla["Promotor"].notna()]
    tabla = tabla[tabla["Año"].notna() & tabla["Mes"].notna()]
    if tabla.empty:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    return (
        tabla.groupby(
            ["Año", "Mes", "Promotor", "Concepto", "Ramo", "Ini_Ren"],
            dropna=False,
            as_index=False,
        )["Monto"]
        .sum()
    )


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
                    "Promotor": df[columnas["PROMOTOR"]].map(_promotor),
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
            return _filas(trabajo.to_dict("records"))
    return pd.DataFrame(columns=COLUMNAS_DATOS)


def extraer_comision_base(ruta, anio):
    registros = []
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
    for hoja in hojas:
        df = pd.read_excel(ruta, sheet_name=hoja)
        if df.empty:
            continue
        columnas = {_norm_nombre(col): col for col in df.columns}
        if "PROMOTORIA" not in columnas or "COMISION" not in columnas:
            continue
        for _, fila in df.iterrows():
            ini = _ini_tabla(fila.get(columnas.get("TIPO COMISION")))
            if not ini:
                continue
            ramo = _ramo(fila.get(columnas.get("TIPO"), ""))
            if not ramo:
                continue
            registros.append(
                {
                    "Año": anio,
                    "Mes": _mes_numero(fila.get(columnas.get("MES"))),
                    "Promotor": _promotor(fila.get(columnas["PROMOTORIA"])),
                    "Concepto": CONCEPTO_COMISION,
                    "Ramo": ramo,
                    "Ini_Ren": ini,
                    "Monto": fila.get(columnas["COMISION"]),
                }
            )
    return _filas(registros)


def extraer_bonos(ruta, anio):
    if not _tiene_hoja(ruta, "Pagos_de_Bonos"):
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    df = pd.read_excel(ruta, sheet_name="Pagos_de_Bonos")
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    columnas = {_norm_nombre(col): col for col in df.columns}
    col_prom = columnas.get("PROMOTORIA")
    col_importe = columnas.get("IMPORTE")
    if not col_prom or not col_importe:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    col_ini = columnas.get("INI_REN")
    col_ramo = columnas.get("RAMO") or columnas.get("RAMO_CLASIF")
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
    registros = []
    for _, fila in df.iterrows():
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
        ramo = _ramo(fila.get(col_ramo) if col_ramo else None)
        if not ramo:
            continue
        anio_fila = anio
        mes = None
        if col_fecha:
            fecha = pd.to_datetime(fila.get(col_fecha), errors="coerce")
            if pd.notna(fecha):
                anio_fila = int(fecha.year)
                mes = int(fecha.month)
        if mes is None and col_mes:
            mes = _mes_numero(fila.get(col_mes))
        registros.append(
            {
                "Año": anio_fila,
                "Mes": mes,
                "Promotor": _promotor(fila.get(col_prom)),
                "Concepto": concepto,
                "Ramo": ramo,
                "Ini_Ren": ini,
                "Monto": fila.get(col_importe),
            }
        )
    return _filas(registros)


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
        registros.append(
            {
                "Año": _a_entero(fila.get(col_anio)),
                "Mes": _mes_numero(fila.get(col_mes)),
                "Promotor": _promotor(fila.get(col_prom)),
                "Concepto": CONCEPTO_ANP,
                "Ramo": ramo,
                "Ini_Ren": None,
                "Monto": fila.get(col_monto),
            }
        )
    return _filas(registros)


def _mapear_titulo_prima(titulo):
    texto = _norm_nombre(titulo)
    for fragmento, ramo, ini in MAPA_PRIMAS:
        if fragmento in texto:
            return ramo, ini
    ramo = "GMM" if "GMM" in texto or "PRIMORDIAL" in texto or "METPLUS" in texto else None
    if "VIDA" in texto:
        ramo = "VIDA"
    if not ramo:
        return None
    if "RENOV" in texto:
        return ramo, INI_RENOV
    return ramo, INI_PRIMER


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
                        "Concepto": CONCEPTO_PRIMA,
                        "Ramo": ramo,
                        "Ini_Ren": ini,
                        "Monto": fila[col_mes],
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
    return _filas(pd.concat(partes, ignore_index=True).to_dict("records"))


def consolidar_datos(paquetes, actualizar_estado=None):
    partes = []
    for anio, ruta in paquetes:
        if not ruta:
            continue
        partes.append(extraer_paquete(ruta, anio, actualizar_estado))
    partes = [parte for parte in partes if not parte.empty]
    if not partes:
        return pd.DataFrame(columns=COLUMNAS_DATOS)
    return _filas(pd.concat(partes, ignore_index=True).to_dict("records"))
