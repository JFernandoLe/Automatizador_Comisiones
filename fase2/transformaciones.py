import pandas as pd

COL_CLAVE = "CLAVE CON CAMBIO DE PF A PM"
COL_PROMOTORIA = "PROMOTORIA"
COL_COMISION = "COMISION"
COL_TIPO_COM = "TIPO COMISIÓN"
COL_ROL = "COMO PROMOTOR / AGENTE"
COL_TIPO_PROD = "Tipo"

ROLES_AGENTE = {"AGENTE", "AGENTES"}
ROLES_PROMOTOR = {"PROMOTOR"}
ROLES_IGNORAR = {"CLAVES DIRECTAS"}

TIPO_1ER = "1ER. AÑO"
TIPO_OTROS = "OTROS CONCEPTOS"
TIPO_RENOV = "RENOVACIÓN"

TIPOS_VIDA_AGENTE = (TIPO_1ER, TIPO_OTROS, TIPO_RENOV)
TIPOS_VIDA_PROMOTOR = (TIPO_1ER, TIPO_RENOV)
TIPOS_GMM_AGENTE = (TIPO_1ER, TIPO_OTROS, TIPO_RENOV)
TIPOS_GMM_PROMOTOR = (TIPO_1ER, TIPO_OTROS, TIPO_RENOV)

COLUMNAS_VIDA = [
    "prom",
    "Promotor",
    "Clave con PF",
    "AGENTE_1ER_META",
    "AGENTE_1ER_VIDA",
    "AGENTE_1ER_TOTAL",
    "AGENTE_OTROS_META",
    "AGENTE_OTROS_VIDA",
    "AGENTE_OTROS_TOTAL",
    "AGENTE_REN_META",
    "AGENTE_REN_VIDA",
    "AGENTE_REN_TOTAL",
    "AGENTE_TOTAL",
    "PROMOTOR_1ER_META",
    "PROMOTOR_1ER_VIDA",
    "PROMOTOR_1ER_TOTAL",
    "PROMOTOR_REN_META",
    "PROMOTOR_REN_VIDA",
    "PROMOTOR_REN_TOTAL",
    "PROMOTOR_TOTAL",
    "TOTAL_GENERAL",
]

COLUMNAS_GMM = [
    "prom",
    "Promotor",
    "Clave con PF a PM",
    "AGENTE_1ER",
    "AGENTE_OTROS",
    "AGENTE_REN",
    "AGENTE_TOTAL",
    "PROMOTOR_1ER",
    "PROMOTOR_OTROS",
    "PROMOTOR_REN",
    "PROMOTOR_TOTAL",
    "TOTAL_GENERAL",
]


def _a_entero(valor):
    try:
        return int(float(str(valor).strip()))
    except (TypeError, ValueError):
        return None


def crear_pfpm_map(df_pfpm):
    mapa = {}
    for original, reporteria in zip(
        df_pfpm["AGENTE_ORIGINAL"], df_pfpm["AGENTE_REPORTERIA"]
    ):
        clave = _a_entero(original)
        valor = _a_entero(reporteria)
        if clave is not None and valor is not None:
            mapa[clave] = valor
    return mapa


def crear_dist_map(df_dist):
    if df_dist.shape[1] < 2:
        raise ValueError(
            "Distribución Comercial debe tener al menos dos columnas (A y B)."
        )
    mapa = {}
    for original, destino in zip(df_dist.iloc[:, 0], df_dist.iloc[:, 1]):
        clave = _a_entero(original)
        if clave is None:
            continue
        valor = _a_entero(destino)
        mapa[clave] = valor if valor is not None else destino
    return mapa


def resolver_prom(promotoria, dist_map, pfpm_map):
    clave = _a_entero(promotoria)
    if clave is None or clave not in dist_map:
        return None
    dist_val = dist_map[clave]
    dist_int = _a_entero(dist_val)
    if dist_int is not None and dist_int in pfpm_map:
        return pfpm_map[dist_int]
    return dist_val


def _normalizar_rol(valor):
    texto = str(valor).strip().upper()
    if texto in ROLES_IGNORAR:
        return None
    if texto in ROLES_AGENTE:
        return "AGENTE"
    if texto in ROLES_PROMOTOR:
        return "PROMOTOR"
    return None


def _normalizar_tipo_com(valor):
    texto = (
        str(valor)
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )
    if "1ER" in texto or "PRIMER" in texto:
        return TIPO_1ER
    if "OTRO" in texto:
        return TIPO_OTROS
    if "RENOV" in texto:
        return TIPO_RENOV
    return None


def _producto(valor):
    if "METALIFE" in str(valor).upper():
        return "META"
    return "VIDA"


def _clave_grupo(valor):
    entero = _a_entero(valor)
    return entero if entero is not None else str(valor).strip()


def _preparar_base(df, con_producto=False):
    faltantes = [
        col
        for col in (COL_CLAVE, COL_PROMOTORIA, COL_COMISION, COL_TIPO_COM, COL_ROL)
        if col not in df.columns
    ]
    if con_producto and COL_TIPO_PROD not in df.columns:
        faltantes.append(COL_TIPO_PROD)
    if faltantes:
        raise ValueError(
            "El reporte no contiene las columnas obligatorias: "
            + ", ".join(faltantes)
        )

    salida = pd.DataFrame()
    salida["PROMOTORIA"] = df[COL_PROMOTORIA].map(_clave_grupo)
    salida["CLAVE"] = df[COL_CLAVE].map(_clave_grupo)
    salida["COMISION"] = pd.to_numeric(df[COL_COMISION], errors="coerce").fillna(0)
    salida["ROL"] = df[COL_ROL].map(_normalizar_rol)
    salida["TIPO_COM"] = df[COL_TIPO_COM].map(_normalizar_tipo_com)
    if con_producto:
        salida["PRODUCTO"] = df[COL_TIPO_PROD].map(_producto)
    salida = salida[salida["ROL"].notna() & salida["TIPO_COM"].notna()].copy()
    salida = salida[
        salida["PROMOTORIA"].astype(str).str.strip().ne("")
        & salida["CLAVE"].astype(str).str.strip().ne("")
    ]
    return salida


def _sumas_vida(base):
    agrupado = (
        base.groupby(
            ["PROMOTORIA", "CLAVE", "ROL", "TIPO_COM", "PRODUCTO"],
            dropna=False,
        )["COMISION"]
        .sum()
    )
    return agrupado


def _sumas_gmm(base):
    return (
        base.groupby(
            ["PROMOTORIA", "CLAVE", "ROL", "TIPO_COM"],
            dropna=False,
        )["COMISION"]
        .sum()
    )


def _pares(base):
    pares = (
        base[["PROMOTORIA", "CLAVE"]]
        .drop_duplicates()
        .sort_values(["PROMOTORIA", "CLAVE"])
    )
    return list(pares.itertuples(index=False, name=None))


def construir_tabla_vida(df, dist_map, pfpm_map):
    base = _preparar_base(df, con_producto=True)
    sumas = _sumas_vida(base)
    filas = []
    for promotoria, clave in _pares(base):
        def monto(rol, tipo, producto):
            clave_suma = (promotoria, clave, rol, tipo, producto)
            if clave_suma in sumas.index:
                return float(sumas.loc[clave_suma])
            return 0.0

        def trio(rol, tipo):
            meta = monto(rol, tipo, "META")
            vida = monto(rol, tipo, "VIDA")
            return meta, vida, meta + vida

        a1_m, a1_v, a1_t = trio("AGENTE", TIPO_1ER)
        ao_m, ao_v, ao_t = trio("AGENTE", TIPO_OTROS)
        ar_m, ar_v, ar_t = trio("AGENTE", TIPO_RENOV)
        agente_total = a1_t + ao_t + ar_t

        p1_m, p1_v, p1_t = trio("PROMOTOR", TIPO_1ER)
        pr_m, pr_v, pr_t = trio("PROMOTOR", TIPO_RENOV)
        promotor_total = p1_t + pr_t

        filas.append(
            {
                "prom": resolver_prom(promotoria, dist_map, pfpm_map),
                "Promotor": promotoria,
                "Clave con PF": clave,
                "AGENTE_1ER_META": a1_m,
                "AGENTE_1ER_VIDA": a1_v,
                "AGENTE_1ER_TOTAL": a1_t,
                "AGENTE_OTROS_META": ao_m,
                "AGENTE_OTROS_VIDA": ao_v,
                "AGENTE_OTROS_TOTAL": ao_t,
                "AGENTE_REN_META": ar_m,
                "AGENTE_REN_VIDA": ar_v,
                "AGENTE_REN_TOTAL": ar_t,
                "AGENTE_TOTAL": agente_total,
                "PROMOTOR_1ER_META": p1_m,
                "PROMOTOR_1ER_VIDA": p1_v,
                "PROMOTOR_1ER_TOTAL": p1_t,
                "PROMOTOR_REN_META": pr_m,
                "PROMOTOR_REN_VIDA": pr_v,
                "PROMOTOR_REN_TOTAL": pr_t,
                "PROMOTOR_TOTAL": promotor_total,
                "TOTAL_GENERAL": agente_total + promotor_total,
            }
        )

    if not filas:
        return pd.DataFrame(columns=COLUMNAS_VIDA)
    return pd.DataFrame(filas)[COLUMNAS_VIDA]


def construir_tabla_gmm(df, dist_map, pfpm_map):
    base = _preparar_base(df, con_producto=False)
    sumas = _sumas_gmm(base)
    filas = []
    for promotoria, clave in _pares(base):
        def monto(rol, tipo):
            clave_suma = (promotoria, clave, rol, tipo)
            if clave_suma in sumas.index:
                return float(sumas.loc[clave_suma])
            return 0.0

        a1 = monto("AGENTE", TIPO_1ER)
        ao = monto("AGENTE", TIPO_OTROS)
        ar = monto("AGENTE", TIPO_RENOV)
        agente_total = a1 + ao + ar
        p1 = monto("PROMOTOR", TIPO_1ER)
        po = monto("PROMOTOR", TIPO_OTROS)
        pr = monto("PROMOTOR", TIPO_RENOV)
        promotor_total = p1 + po + pr
        filas.append(
            {
                "prom": resolver_prom(promotoria, dist_map, pfpm_map),
                "Promotor": promotoria,
                "Clave con PF a PM": clave,
                "AGENTE_1ER": a1,
                "AGENTE_OTROS": ao,
                "AGENTE_REN": ar,
                "AGENTE_TOTAL": agente_total,
                "PROMOTOR_1ER": p1,
                "PROMOTOR_OTROS": po,
                "PROMOTOR_REN": pr,
                "PROMOTOR_TOTAL": promotor_total,
                "TOTAL_GENERAL": agente_total + promotor_total,
            }
        )

    if not filas:
        return pd.DataFrame(columns=COLUMNAS_GMM)
    return pd.DataFrame(filas)[COLUMNAS_GMM]


def _norm_nombre(valor):
    return (
        str(valor)
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )


def _buscar_columna(df, nombre):
    objetivo = _norm_nombre(nombre)
    for col in df.columns:
        if _norm_nombre(col) == objetivo:
            return col
    raise ValueError(f"El archivo Bonos no contiene la columna '{nombre}'.")


def construir_pagos_bonos(df, dist_map, pfpm_map):
    salida = df.copy()
    col_prom = _buscar_columna(salida, "Promotoria")
    col_prima = _buscar_columna(salida, "Prima")
    col_poliza = _buscar_columna(salida, "Póliza")
    original_prom = salida[col_prom].copy()

    def lookup_dist(valor):
        clave = _a_entero(valor)
        if clave is None or clave not in dist_map:
            return None
        return dist_map[clave]

    def lookup_pfpm(valor):
        if valor is None:
            return None
        try:
            if pd.isna(valor):
                return None
        except (TypeError, ValueError):
            pass
        clave = _a_entero(valor)
        if clave is not None and clave in pfpm_map:
            return pfpm_map[clave]
        return valor

    polizas = original_prom.map(lookup_dist)
    salida[col_prima] = original_prom
    salida[col_poliza] = polizas
    salida[col_prom] = polizas.map(lookup_pfpm)
    return salida
