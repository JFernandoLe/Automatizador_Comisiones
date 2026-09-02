import pandas as pd

COLUMNAS_VIDA = [
    "CLAVE CON CAMBIO DE PF A PM", "FECHA", "MES", "AGENTE", "NOMBREAGENTE",
    "INTCONCEPTO", "DESCCONCEPTO", "RAMO", "POLIZA", "PRIMA", "COMISION",
    "USUARIO", "INTCONSECUTIVO", "PROMOTORIA", "DESCPROMOTORIA", "TIPOPERSONA",
    "CP", "DESCPRESP", "PROMOTORIA_1", "DESCPROMOTORIA_1", "TIPO COMISIÓN",
    "COMO PROMOTOR / AGENTE", "Cobertura", "Tipo"
]

COLUMNAS_GMM = [
    "CLAVE CON CAMBIO DE PF A PM", "FECHA", "MES", "AGENTE", "NOMBREAGENTE",
    "INTCONCEPTO", "DESCCONCEPTO", "RAMO", "POLIZA", "PRIMA", "COMISION",
    "USUARIO", "INTCONSECUTIVO", "PROMOTORIA", "DESCPROMOTORIA", "TIPOPERSONA",
    "CP", "DESCPRESP", "PROMOTORIA_1", "DESCPROMOTORIA_1", "TIPO COMISIÓN",
    "COMO PROMOTOR / AGENTE"
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


def crear_catalogo_conceptos(df_conceptos):
    catalogo = {}
    for _, fila in df_conceptos.iterrows():
        concepto = _a_entero(fila["INTCONCEPTO"])
        if concepto is None:
            continue
        catalogo[concepto] = {
            "TIPO_COMISION": str(fila["TIPO_COMISION"]).strip()
            if pd.notna(fila["TIPO_COMISION"]) else "",
            "COMO_PROMOTOR_AGENTE": str(fila["COMO_PROMOTOR_AGENTE"]).strip()
            if pd.notna(fila["COMO_PROMOTOR_AGENTE"]) else "",
        }
    return catalogo


def obtener_clave_pf_pm(agente, pfpm_map):
    agente_num = _a_entero(agente)
    if agente_num is None:
        return agente
    return pfpm_map.get(agente_num, agente_num)


def convertir_tipos(df):
    enteros = [
        "CLAVE CON CAMBIO DE PF A PM", "AGENTE", "INTCONCEPTO", "RAMO",
        "INTCONSECUTIVO", "PROMOTORIA", "PROMOTORIA_1"
    ]
    decimales = ["PRIMA", "COMISION"]

    for col in enteros:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    for col in decimales:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")

    for col in ["POLIZA", "CP"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    for col in df.columns:
        if col not in enteros + decimales + ["FECHA", "POLIZA", "CP"]:
            df[col] = df[col].fillna("").astype(str).str.strip()
    return df


def obtener_tipo_comision(intconcepto, catalogo):
    concepto = _a_entero(intconcepto)
    if concepto is None:
        return ""
    return catalogo.get(concepto, {}).get("TIPO_COMISION", "")


def obtener_como_promotor_agente(usuario, intconcepto, catalogo):
    usuario = str(usuario).strip().upper()
    if usuario in ["AGENTE", "AGENTES"]:
        return "AGENTE"
    if usuario == "PROMOTOR":
        return "PROMOTOR"
    concepto = _a_entero(intconcepto)
    if concepto is None:
        return ""
    return catalogo.get(concepto, {}).get("COMO_PROMOTOR_AGENTE", "")


def construir_vida_etapa4(df, df_vlsp, df_tipo, df_pfpm, df_conceptos_vida):
    df = df.copy()
    df_vlsp = df_vlsp.copy()
    pfpm_map = crear_pfpm_map(df_pfpm)
    conceptos_vida = crear_catalogo_conceptos(df_conceptos_vida)

    df["POLIZA"] = df["POLIZA"].fillna("").astype(str).str.strip()
    df_vlsp["POLIZA"] = df_vlsp["POLIZA"].fillna("").astype(str).str.strip()
    cobertura_map = dict(zip(df_vlsp["POLIZA"], df_vlsp["COBERTURA"]))
    df["Cobertura"] = df["POLIZA"].map(cobertura_map)

    tipo_map = dict(zip(df_tipo["Clave Cobertura"], df_tipo["Producto"]))

    def obtener_tipo(cobertura):
        producto = tipo_map.get(cobertura, "")
        return "METALIFE" if "METALIFE" in str(producto).upper() else "VIDA"

    df["Tipo"] = df["Cobertura"].apply(obtener_tipo)
    df["CLAVE CON CAMBIO DE PF A PM"] = df["AGENTE"].apply(
        lambda agente: obtener_clave_pf_pm(agente, pfpm_map)
    )
    df["TIPO COMISIÓN"] = df["INTCONCEPTO"].apply(
        lambda concepto: obtener_tipo_comision(concepto, conceptos_vida)
    )
    df["COMO PROMOTOR / AGENTE"] = df.apply(
        lambda fila: obtener_como_promotor_agente(
            fila["USUARIO"], fila["INTCONCEPTO"], conceptos_vida
        ), axis=1
    )
    return convertir_tipos(df[COLUMNAS_VIDA].copy())


def construir_gmm_etapa4(df, df_pfpm, df_conceptos_gmm):
    df = df.copy()
    pfpm_map = crear_pfpm_map(df_pfpm)
    conceptos_gmm = crear_catalogo_conceptos(df_conceptos_gmm)

    df["CLAVE CON CAMBIO DE PF A PM"] = df["AGENTE"].apply(
        lambda agente: obtener_clave_pf_pm(agente, pfpm_map)
    )
    df["TIPO COMISIÓN"] = df["INTCONCEPTO"].apply(
        lambda concepto: obtener_tipo_comision(concepto, conceptos_gmm)
    )
    df["COMO PROMOTOR / AGENTE"] = df.apply(
        lambda fila: obtener_como_promotor_agente(
            fila["USUARIO"], fila["INTCONCEPTO"], conceptos_gmm
        ), axis=1
    )
    return convertir_tipos(df[COLUMNAS_GMM].copy())
