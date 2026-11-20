import shutil

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from ingresos.fuentes import (
    CONCEPTO_ANP,
    CONCEPTO_APOYO,
    CONCEPTO_BONO,
    CONCEPTO_COMISION,
    CONCEPTO_CONVENIO,
    CONCEPTO_PRIMA,
    COLUMNAS_DATOS,
    INI_PRIMER,
    INI_RENOV,
)

HOJA_TABLERO = "Filtro Ingresos vs Ventas"
HOJA_PROMOTOR = "Filtro Promotor Ingresos"
MESES_NOMBRE = (
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
)

COLUMNAS_MES_H1 = (
    ("D", 1, -1),
    ("E", 2, -1),
    ("F", 3, -1),
    ("H", 1, 0),
    ("I", 2, 0),
    ("J", 3, 0),
    ("N", 4, -1),
    ("O", 5, -1),
    ("P", 6, -1),
    ("R", 4, 0),
    ("S", 5, 0),
    ("T", 6, 0),
    ("X", 7, -1),
    ("Y", 8, -1),
    ("Z", 9, -1),
    ("AB", 7, 0),
    ("AC", 8, 0),
    ("AD", 9, 0),
)

HOJA1_HOJAS = {
    10: (CONCEPTO_ANP, "VIDA", None),
    11: (CONCEPTO_ANP, "GMM", None),
    15: (CONCEPTO_PRIMA, "VIDA", INI_PRIMER),
    16: (CONCEPTO_PRIMA, "VIDA", INI_RENOV),
    18: (CONCEPTO_PRIMA, "GMM", INI_PRIMER),
    19: (CONCEPTO_PRIMA, "GMM", INI_RENOV),
    27: (CONCEPTO_COMISION, "VIDA", INI_PRIMER),
    28: (CONCEPTO_COMISION, "VIDA", INI_RENOV),
    30: (CONCEPTO_COMISION, "GMM", INI_PRIMER),
    31: (CONCEPTO_COMISION, "GMM", INI_RENOV),
    38: (CONCEPTO_APOYO, "VIDA", None),
    39: (CONCEPTO_APOYO, "GMM", None),
    42: (CONCEPTO_CONVENIO, "VIDA", None),
    43: (CONCEPTO_CONVENIO, "GMM", None),
    47: (CONCEPTO_BONO, "VIDA", INI_PRIMER),
    48: (CONCEPTO_BONO, "VIDA", INI_RENOV),
    50: (CONCEPTO_BONO, "GMM", INI_PRIMER),
    51: (CONCEPTO_BONO, "GMM", INI_RENOV),
}

HOJA1_PADRES = {
    9: (10, 11),
    14: (15, 16),
    17: (18, 19),
    13: (14, 17),
    26: (27, 28),
    29: (30, 31),
    25: (26, 29),
    37: (38, 39),
    41: (42, 43),
    46: (47, 48),
    49: (50, 51),
    45: (46, 49),
    34: (38, 42, 46),
    35: (39, 43, 49),
    33: (34, 35),
    22: (26, 34),
    23: (29, 35),
    21: (22, 23),
}

HOJA1_YTD = (("AH", -1), ("AI", 0))

# (fila, concepto, ramo, ini, prima_fila)
HOJA2_BLOQUES = (
    {
        "anio_ant2": "E",
        "anio_ant1": "F",
        "ytd_ant1": "M",
        "ytd_act": "N",
        "ratio_ant2": "J",
        "ratio_ant1": "K",
        "ratio_ytd_ant1": "R",
        "ratio_ytd_act": "S",
        "filas": {
            9: (CONCEPTO_COMISION, None, None, 14),
            10: (CONCEPTO_BONO, None, None, 14),
            11: (CONCEPTO_APOYO, None, None, 14),
            14: (CONCEPTO_PRIMA, None, None, None),
            17: (CONCEPTO_COMISION, None, INI_PRIMER, 21),
            18: (CONCEPTO_BONO, None, INI_PRIMER, 21),
            21: (CONCEPTO_PRIMA, None, INI_PRIMER, None),
            24: (CONCEPTO_COMISION, None, INI_RENOV, 28),
            25: (CONCEPTO_BONO, None, INI_RENOV, 28),
            28: (CONCEPTO_PRIMA, None, INI_RENOV, None),
        },
        "padres": {12: (9, 10, 11), 19: (17, 18), 26: (24, 25)},
    },
    {
        "anio_ant2": "V",
        "anio_ant1": "W",
        "ytd_ant1": "AD",
        "ytd_act": "AE",
        "ratio_ant2": "AA",
        "ratio_ant1": "AB",
        "ratio_ytd_ant1": "AI",
        "ratio_ytd_act": "AJ",
        "filas": {
            9: (CONCEPTO_COMISION, "VIDA", None, 14),
            10: (CONCEPTO_BONO, "VIDA", None, 14),
            11: (CONCEPTO_APOYO, "VIDA", None, 14),
            14: (CONCEPTO_PRIMA, "VIDA", None, None),
            17: (CONCEPTO_COMISION, "VIDA", INI_PRIMER, 21),
            18: (CONCEPTO_BONO, "VIDA", INI_PRIMER, 21),
            21: (CONCEPTO_PRIMA, "VIDA", INI_PRIMER, None),
            24: (CONCEPTO_COMISION, "VIDA", INI_RENOV, 28),
            25: (CONCEPTO_BONO, "VIDA", INI_RENOV, 28),
            28: (CONCEPTO_PRIMA, "VIDA", INI_RENOV, None),
        },
        "padres": {12: (9, 10, 11), 19: (17, 18), 26: (24, 25)},
    },
    {
        "anio_ant2": "AM",
        "anio_ant1": "AN",
        "ytd_ant1": "AU",
        "ytd_act": "AV",
        "ratio_ant2": "AR",
        "ratio_ant1": "AS",
        "ratio_ytd_ant1": "AZ",
        "ratio_ytd_act": "BA",
        "filas": {
            9: (CONCEPTO_COMISION, "GMM", None, 14),
            10: (CONCEPTO_BONO, "GMM", None, 14),
            11: (CONCEPTO_APOYO, "GMM", None, 14),
            14: (CONCEPTO_PRIMA, "GMM", None, None),
            17: (CONCEPTO_COMISION, "GMM", INI_PRIMER, 21),
            18: (CONCEPTO_BONO, "GMM", INI_PRIMER, 21),
            21: (CONCEPTO_PRIMA, "GMM", INI_PRIMER, None),
            24: (CONCEPTO_COMISION, "GMM", INI_RENOV, 28),
            25: (CONCEPTO_BONO, "GMM", INI_RENOV, 28),
            28: (CONCEPTO_PRIMA, "GMM", INI_RENOV, None),
        },
        "padres": {12: (9, 10, 11), 19: (17, 18), 26: (24, 25)},
    },
)

PROMOTOR_CELDA = f"'{HOJA_TABLERO}'!$C$3"


def _formula_suma(anio_expr, mes_expr, concepto, ramo, ini_ren):
    partes = [
        "Datos[Monto]",
        f"Datos[Año],{anio_expr}",
    ]
    if mes_expr:
        partes.append(f"Datos[Mes],{mes_expr}")
    partes.append(f'Datos[Concepto],"{concepto}"')
    if ramo:
        partes.append(f'Datos[Ramo],"{ramo}"')
    if ini_ren:
        partes.append(f'Datos[Ini_Ren],"{ini_ren}"')
    filtros = ",".join(partes)
    filtros_prom = f"{filtros},Datos[Promotoria],{PROMOTOR_CELDA}"
    cuerpo = (
        f'IF({PROMOTOR_CELDA}="TOTAL",SUMIFS({filtros}),SUMIFS({filtros_prom}))'
    )
    return f"=IFERROR({cuerpo}/1000000,0)"


def _anio_param(offset):
    if offset == 0:
        return "Parametros!$B$1"
    if offset == -1:
        return "Parametros!$B$4"
    return "Parametros!$B$5"


def _escribir_parametros(wb, anio, mes_corte):
    if "Parametros" in wb.sheetnames:
        del wb["Parametros"]
    ws = wb.create_sheet("Parametros", 0)
    ws["A1"] = "Año actual"
    ws["B1"] = int(anio)
    ws["A2"] = "Mes corte"
    ws["B2"] = int(mes_corte)
    ws["A3"] = "Promotor (se elige en la hoja de filtro)"
    ws["A4"] = "Año anterior"
    ws["B4"] = "=B1-1"
    ws["A5"] = "Hace dos años"
    ws["B5"] = "=B1-2"
    ws["C1"] = "Cambie el año o el mes de corte; las tablas se recalculan."
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 54


def _escribir_promotores(wb, claves):
    if "Promotores" in wb.sheetnames:
        del wb["Promotores"]
    ws = wb.create_sheet("Promotores", 1)
    ws["A1"] = "Clave"
    ws["B1"] = "Nombre"
    ws["A2"] = "TOTAL"
    ws["B2"] = "Todos los promotores"
    unicos = []
    vistos = set()
    for clave in claves:
        if clave in vistos or clave is None:
            continue
        vistos.add(clave)
        unicos.append(clave)
    for offset, clave in enumerate(sorted(unicos, key=lambda valor: (str(type(valor)), str(valor)))):
        ws.cell(row=3 + offset, column=1, value=clave)
    ultima = max(2, 2 + len(unicos))
    tabla = Table(displayName="Promotores", ref=f"A1:B{ultima}")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabla)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28
    return ultima


def _escribir_datos(wb, df):
    if "Datos" in wb.sheetnames:
        del wb["Datos"]
    ws = wb.create_sheet("Datos", 2)
    ws.append(list(COLUMNAS_DATOS))
    if df is None or df.empty:
        ws.append([0, 1, None, "TOTAL", None, "TOTAL", "ANP", "VIDA", None, 0])
        ultima = 2
    else:
        trabajo = df[COLUMNAS_DATOS].copy()
        for fila in trabajo.itertuples(index=False, name=None):
            ws.append([None if pd_isna(valor) else valor for valor in fila])
        ultima = 1 + len(trabajo)
    letra = get_column_letter(len(COLUMNAS_DATOS))
    tabla = Table(displayName="Datos", ref=f"A1:{letra}{ultima}")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tabla)
    anchos = (10, 8, 12, 14, 14, 14, 14, 10, 16, 16)
    for letra_col, ancho in zip("ABCDEFGHIJ", anchos):
        ws.column_dimensions[letra_col].width = ancho


def pd_isna(valor):
    if valor is None:
        return True
    try:
        import pandas as pd

        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _aplicar_tablero(ws, anio, mes_corte):
    lista = ",".join(f'"{nombre}"' for nombre in MESES_NOMBRE)
    ws["B1"] = f'=CONCATENATE("YTD ",CHOOSE(Parametros!$B$2,{lista}))'
    ws["C3"] = "TOTAL"
    ws["C4"] = '=IF(C3="TOTAL","Todos los promotores",IFERROR(VLOOKUP(C3,Promotores!A:B,2,FALSE),""))'
    ws["G8"] = "=Parametros!$B$4"
    ws["K8"] = "=Parametros!$B$1"
    ws["Q8"] = "=Parametros!$B$4"
    ws["U8"] = "=Parametros!$B$1"
    ws["AA8"] = "=Parametros!$B$4"
    ws["AE8"] = "=Parametros!$B$1"
    ws["AH8"] = "=Parametros!$B$4"
    ws["AI8"] = "=Parametros!$B$1"

    for fila, (concepto, ramo, ini) in HOJA1_HOJAS.items():
        for col, mes, offset in COLUMNAS_MES_H1:
            ws[f"{col}{fila}"] = _formula_suma(
                _anio_param(offset), mes, concepto, ramo, ini
            )
            ws[f"{col}{fila}"].number_format = "#,##0.00"
        for col, offset in HOJA1_YTD:
            ws[f"{col}{fila}"] = _formula_suma(
                _anio_param(offset),
                '"<="&Parametros!$B$2',
                concepto,
                ramo,
                ini,
            )
            ws[f"{col}{fila}"].number_format = "#,##0.00"

    columnas_padre = [col for col, _mes, _off in COLUMNAS_MES_H1] + ["AH", "AI"]
    for fila, hijos in HOJA1_PADRES.items():
        for col in columnas_padre:
            ws[f"{col}{fila}"] = "=" + "+".join(f"{col}{hijo}" for hijo in hijos)
            ws[f"{col}{fila}"].number_format = "#,##0.00"

    dv = DataValidation(
        type="list",
        formula1='INDIRECT("Promotores[Clave]")',
        allow_blank=False,
        showDropDown=False,
    )
    dv.error = "Seleccione un promotor de la lista o TOTAL."
    dv.errorTitle = "Promotor"
    dv.prompt = "Promotor o TOTAL"
    dv.promptTitle = "Filtro"
    ws.add_data_validation(dv)
    dv.add("C3")


def _aplicar_promotor(ws):
    lista = ",".join(f'"{nombre}"' for nombre in MESES_NOMBRE)
    ws["B1"] = f"=CONCATENATE(\"YTD \",CHOOSE(Parametros!$B$2,{lista}))"
    ws["C3"] = f"='{HOJA_TABLERO}'!C3"
    ws["C4"] = f"='{HOJA_TABLERO}'!C4"
    encabezados = (
        ("E8", "Parametros!$B$5"),
        ("F8", "Parametros!$B$4"),
        ("V8", "Parametros!$B$5"),
        ("W8", "Parametros!$B$4"),
        ("AM8", "Parametros!$B$5"),
        ("AN8", "Parametros!$B$4"),
        ("E16", "Parametros!$B$5"),
        ("F16", "Parametros!$B$4"),
        ("V16", "Parametros!$B$5"),
        ("W16", "Parametros!$B$4"),
        ("AM16", "Parametros!$B$5"),
        ("AN16", "Parametros!$B$4"),
        ("E23", "Parametros!$B$5"),
        ("F23", "Parametros!$B$4"),
        ("V23", "Parametros!$B$5"),
        ("W23", "Parametros!$B$4"),
        ("AM23", "Parametros!$B$5"),
        ("AN23", "Parametros!$B$4"),
    )
    for celda, formula in encabezados:
        ws[celda] = f"={formula}"
    for prefijo, anio_celda in (
        ("M", "Parametros!$B$4"),
        ("N", "Parametros!$B$1"),
        ("AD", "Parametros!$B$4"),
        ("AE", "Parametros!$B$1"),
        ("AU", "Parametros!$B$4"),
        ("AV", "Parametros!$B$1"),
    ):
        for fila in (8, 16, 23):
            ws[f"{prefijo}{fila}"] = f'=CONCATENATE("YTD ",{anio_celda})'

    for bloque in HOJA2_BLOQUES:
        for fila, (concepto, ramo, ini, prima_fila) in bloque["filas"].items():
            ws[f"{bloque['anio_ant2']}{fila}"] = _formula_suma(
                _anio_param(-2), None, concepto, ramo, ini
            )
            ws[f"{bloque['anio_ant1']}{fila}"] = _formula_suma(
                _anio_param(-1), None, concepto, ramo, ini
            )
            ws[f"{bloque['ytd_ant1']}{fila}"] = _formula_suma(
                _anio_param(-1), '"<="&Parametros!$B$2', concepto, ramo, ini
            )
            ws[f"{bloque['ytd_act']}{fila}"] = _formula_suma(
                _anio_param(0), '"<="&Parametros!$B$2', concepto, ramo, ini
            )
            for col in (
                bloque["anio_ant2"],
                bloque["anio_ant1"],
                bloque["ytd_ant1"],
                bloque["ytd_act"],
            ):
                ws[f"{col}{fila}"].number_format = "#,##0.00"
            if prima_fila:
                ws[f"{bloque['ratio_ant2']}{fila}"] = (
                    f"=IFERROR({bloque['anio_ant2']}{fila}/{bloque['anio_ant2']}{prima_fila},0)"
                )
                ws[f"{bloque['ratio_ant1']}{fila}"] = (
                    f"=IFERROR({bloque['anio_ant1']}{fila}/{bloque['anio_ant1']}{prima_fila},0)"
                )
                ws[f"{bloque['ratio_ytd_ant1']}{fila}"] = (
                    f"=IFERROR({bloque['ytd_ant1']}{fila}/{bloque['ytd_ant1']}{prima_fila},0)"
                )
                ws[f"{bloque['ratio_ytd_act']}{fila}"] = (
                    f"=IFERROR({bloque['ytd_act']}{fila}/{bloque['ytd_act']}{prima_fila},0)"
                )
                for col in (
                    bloque["ratio_ant2"],
                    bloque["ratio_ant1"],
                    bloque["ratio_ytd_ant1"],
                    bloque["ratio_ytd_act"],
                ):
                    ws[f"{col}{fila}"].number_format = "0.0%"
        for fila, hijos in bloque["padres"].items():
            ramo_padre = bloque["filas"][hijos[0]][1]
            for col, anio_off, mes_expr in (
                (bloque["anio_ant2"], -2, None),
                (bloque["anio_ant1"], -1, None),
                (bloque["ytd_ant1"], -1, '"<="&Parametros!$B$2'),
                (bloque["ytd_act"], 0, '"<="&Parametros!$B$2'),
            ):
                suma_hijos = "+".join(f"{col}{hijo}" for hijo in hijos)
                if fila == 12:
                    extra = _formula_suma(
                        _anio_param(anio_off),
                        mes_expr,
                        CONCEPTO_CONVENIO,
                        ramo_padre,
                        None,
                    )[1:]
                    ws[f"{col}{fila}"] = f"={suma_hijos}+{extra}"
                else:
                    ws[f"{col}{fila}"] = "=" + suma_hijos
                ws[f"{col}{fila}"].number_format = "#,##0.00"
            prima = 14 if fila == 12 else (21 if fila == 19 else 28)
            ws[f"{bloque['ratio_ant2']}{fila}"] = (
                f"=IFERROR({bloque['anio_ant2']}{fila}/{bloque['anio_ant2']}{prima},0)"
            )
            ws[f"{bloque['ratio_ant1']}{fila}"] = (
                f"=IFERROR({bloque['anio_ant1']}{fila}/{bloque['anio_ant1']}{prima},0)"
            )
            ws[f"{bloque['ratio_ytd_ant1']}{fila}"] = (
                f"=IFERROR({bloque['ytd_ant1']}{fila}/{bloque['ytd_ant1']}{prima},0)"
            )
            ws[f"{bloque['ratio_ytd_act']}{fila}"] = (
                f"=IFERROR({bloque['ytd_act']}{fila}/{bloque['ytd_act']}{prima},0)"
            )
            for col in (
                bloque["ratio_ant2"],
                bloque["ratio_ant1"],
                bloque["ratio_ytd_ant1"],
                bloque["ratio_ytd_act"],
            ):
                ws[f"{col}{fila}"].number_format = "0.0%"


def llenar_plantilla(ruta_plantilla, ruta_salida, df_datos, anio, mes_corte):
    shutil.copyfile(ruta_plantilla, ruta_salida)
    wb = load_workbook(ruta_salida)
    _escribir_parametros(wb, anio, mes_corte)
    claves = []
    if df_datos is not None and not df_datos.empty:
        claves = list(df_datos["Promotoria"].dropna().unique())
        if not claves:
            claves = list(df_datos["Promotor"].dropna().unique())
    _escribir_promotores(wb, claves)
    _escribir_datos(wb, df_datos)
    if HOJA_TABLERO not in wb.sheetnames or HOJA_PROMOTOR not in wb.sheetnames:
        raise ValueError(
            "La plantilla no contiene las hojas Filtro Ingresos vs Ventas "
            "y Filtro Promotor Ingresos."
        )
    _aplicar_tablero(wb[HOJA_TABLERO], anio, mes_corte)
    _aplicar_promotor(wb[HOJA_PROMOTOR])
    for nombre in (HOJA_TABLERO, HOJA_PROMOTOR, "Parametros", "Promotores", "Datos"):
        if nombre in wb.sheetnames:
            wb[nombre].sheet_view.showGridLines = nombre in {"Parametros", "Promotores", "Datos"}
    wb.active = wb.sheetnames.index(HOJA_TABLERO)
    wb.save(ruta_salida)
    return str(ruta_salida)
