from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

import pandas as pd

FUENTE = Font(name="Calibri", size=12)
FUENTE_TITULO = Font(name="Calibri", size=12, bold=True)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
FORMATO_MONTO = "#,##0.00"


def _valor_celda(valor):
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    return valor


def _escribir_tabla(ws, tabla, fila_inicio):
    fila = fila_inicio
    anio = ws.cell(row=fila, column=1, value=tabla["anio"])
    titulo = ws.cell(row=fila, column=2, value=tabla["titulo"])
    anio.font = FUENTE_TITULO
    titulo.font = FUENTE_TITULO
    fila += 1

    encabezados = ["Promotor", "Agente"] + list(range(1, tabla["mes_hasta"] + 1))
    encabezados.append("Total general")
    for indice, valor in enumerate(encabezados, start=2):
        celda = ws.cell(row=fila, column=indice, value=valor)
        celda.font = FUENTE_TITULO
        celda.alignment = CENTRO
    fila += 1

    for registro in tabla["filas"]:
        ws.cell(
            row=fila, column=2, value=_valor_celda(registro["promotor"])
        ).font = FUENTE
        ws.cell(
            row=fila, column=3, value=_valor_celda(registro["agente"])
        ).font = FUENTE
        for mes in range(1, tabla["mes_hasta"] + 1):
            celda = ws.cell(
                row=fila,
                column=3 + mes,
                value=_valor_celda(registro["meses"].get(mes) or None),
            )
            celda.font = FUENTE
            celda.number_format = FORMATO_MONTO
        total = ws.cell(
            row=fila,
            column=4 + tabla["mes_hasta"],
            value=_valor_celda(registro["total"]),
        )
        total.font = FUENTE
        total.number_format = FORMATO_MONTO
        fila += 1
    return fila + 2


def guardar_primas(tablas, ruta="Primas.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bases PP"
    fila = 1
    mes_max = 1
    for tabla in tablas:
        mes_max = max(mes_max, tabla["mes_hasta"])
        fila = _escribir_tabla(ws, tabla, fila)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    for mes in range(1, mes_max + 2):
        ws.column_dimensions[get_column_letter(3 + mes)].width = 14
    wb.save(ruta)
    return ruta
