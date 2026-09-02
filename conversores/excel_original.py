"""
Adaptación de ConvExc para un solo archivo y hojas seleccionadas.

Conserva: búsqueda de encabezado Fecha/Agente, salto de la fila posterior
al encabezado en la primera hoja del libro, filtros, reglas DESPAGO y
partición por límite de Excel.
"""
import os
import tempfile
from pathlib import Path

from openpyxl import Workbook
from python_calamine import CalamineWorkbook

from conversores.config.config_loader import cargar_configuracion
from conversores.reglas import procesar_fila

LIMITE_XLSX = 1_048_576
MAX_DATOS_XLSX = LIMITE_XLSX - 1


def encontrar_encabezado(rows):
    """
    Busca la fila que contiene Fecha y Agente.
    """
    for idx, row in enumerate(rows):
        row_str = [str(x).strip() for x in row]

        if "Fecha" in row_str and "Agente" in row_str:
            return idx

    raise Exception("No se encontró la fila de encabezado")


def _crear_xlsx(ruta, encabezado):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Datos")
    ws.append(encabezado)
    return wb, ws, ruta


def _ruta_temporal(prefijo):
    fd, ruta = tempfile.mkstemp(prefix=prefijo, suffix=".xlsx")
    os.close(fd)
    return ruta


def convertir_excel_original(archivo, hojas_seleccionadas, prefijo="convertido_"):
    """
    Convierte un Excel original (.xls / .xlsx) con la lógica de ConvExc
    y deja un .xlsx listo para la lógica existente del automatizador.

    Solo procesa las hojas indicadas, en el orden del libro.
    """
    if not hojas_seleccionadas:
        raise ValueError("Debe seleccionar al menos una hoja.")

    config = cargar_configuracion()
    archivo = Path(archivo)
    wb_calamine = CalamineWorkbook.from_path(str(archivo))
    hojas = list(wb_calamine.sheet_names)
    hojas_objetivo = [h for h in hojas if h in set(hojas_seleccionadas)]

    if not hojas_objetivo:
        print(f"Sin hojas seleccionadas en: {archivo.name}")
        return None

    sheet = wb_calamine.get_sheet_by_name(hojas[0])
    rows = sheet.to_python()
    fila_encabezado = encontrar_encabezado(rows)
    encabezado = rows[fila_encabezado]
    mapa_columnas = {
        str(nombre).strip(): indice
        for indice, nombre in enumerate(encabezado)
        if str(nombre).strip()
    }

    estadisticas = {
        "filas_leidas": 0,
        "filas_eliminadas": 0,
        "filas_modificadas": 0,
        "filas_exportadas": 0,
    }

    ruta_salida = _ruta_temporal(prefijo)
    parte = 1
    filas_actuales = 0
    rutas = []

    wb_out, ws_out, ruta_actual = _crear_xlsx(ruta_salida, encabezado)

    for nombre_hoja in hojas_objetivo:
        indice_hoja = hojas.index(nombre_hoja)
        sheet = wb_calamine.get_sheet_by_name(nombre_hoja)
        rows = sheet.to_python()

        if indice_hoja == 0:
            inicio = fila_encabezado + 2
        else:
            inicio = 0

        for fila in rows[inicio:]:
            fila = list(fila)
            fila = procesar_fila(
                fila,
                config,
                mapa_columnas,
                estadisticas,
            )

            if fila is None:
                continue

            if filas_actuales >= MAX_DATOS_XLSX:
                wb_out.save(ruta_actual)
                rutas.append(ruta_actual)
                parte += 1
                filas_actuales = 0
                ruta_actual = _ruta_temporal(f"{prefijo}p{parte}_")
                wb_out, ws_out, ruta_actual = _crear_xlsx(ruta_actual, encabezado)

            ws_out.append(fila)
            estadisticas["filas_exportadas"] += 1
            filas_actuales += 1

    wb_out.save(ruta_actual)
    rutas.append(ruta_actual)

    print("\nResumen conversión Excel original:")
    print(f"Filas leídas:      {estadisticas['filas_leidas']:,}")
    print(f"Filas eliminadas:  {estadisticas['filas_eliminadas']:,}")
    print(f"Filas modificadas: {estadisticas['filas_modificadas']:,}")
    print(f"Filas exportadas:  {estadisticas['filas_exportadas']:,}")

    if len(rutas) == 1:
        return rutas[0]

    return _consolidar_partes(rutas, encabezado, prefijo)


def convertir_exceles_originales(
    archivos,
    hojas_seleccionadas,
    prefijo="convertido_",
    actualizar_estado=None,
):
    """
    Convierte y consolida varios Excel originales, como ConvExc + Consolidado.xlsx.
    """
    if not archivos:
        raise ValueError("Debe seleccionar al menos un archivo Excel.")
    if not hojas_seleccionadas:
        raise ValueError("Debe seleccionar al menos una hoja.")

    convertidos = []
    omitidos = []
    total = len(archivos)

    for indice, archivo in enumerate(archivos, start=1):
        nombre = Path(archivo).name
        mensaje = f"Convirtiendo {indice}/{total}: {nombre}"
        print(f"\n[{indice}/{total}] {nombre}")
        if actualizar_estado:
            progreso = min(90, int(indice / total * 80))
            actualizar_estado(mensaje, progreso)

        try:
            ruta = convertir_excel_original(
                archivo,
                hojas_seleccionadas,
                prefijo=f"{prefijo}{indice}_",
            )
            if ruta:
                convertidos.append(ruta)
            else:
                omitidos.append(nombre)
        except Exception as error:
            print(f"ERROR en {nombre}: {error}")
            omitidos.append(f"{nombre} ({error})")

    if not convertidos:
        detalle = "; ".join(omitidos[:8])
        raise ValueError(
            "Ningún Excel se pudo convertir. "
            f"Revisados: {total}. Ejemplos: {detalle}"
        )

    print(
        f"\nArchivos convertidos: {len(convertidos)} / {total}. "
        f"Omitidos: {len(omitidos)}"
    )
    if actualizar_estado:
        actualizar_estado("Consolidando archivos convertidos...", 92)

    consolidado = _consolidar_convertidos(convertidos, prefijo)

    for ruta in convertidos:
        try:
            os.remove(ruta)
        except OSError:
            pass

    return consolidado


def _consolidar_convertidos(rutas, prefijo):
    """Une varios xlsx convertidos en un solo libro, como generar_consolidado de ConvExc."""
    from openpyxl import load_workbook

    ruta_final = _ruta_temporal(f"{prefijo}cons_")
    wb_final = Workbook(write_only=True)
    ws = wb_final.create_sheet("Datos_1")
    encabezado = None
    filas_en_hoja = 0
    numero_hoja = 1

    for ruta in rutas:
        wb = load_workbook(ruta, read_only=True, data_only=True)
        for hoja in wb.worksheets:
            primera_fila = True
            for fila in hoja.iter_rows(values_only=True):
                valores = list(fila)
                if primera_fila:
                    primera_fila = False
                    if encabezado is None:
                        encabezado = valores
                        ws.append(encabezado)
                        filas_en_hoja = 1
                    continue

                if filas_en_hoja >= MAX_DATOS_XLSX:
                    numero_hoja += 1
                    ws = wb_final.create_sheet(f"Datos_{numero_hoja}")
                    ws.append(encabezado)
                    filas_en_hoja = 1

                ws.append(valores)
                filas_en_hoja += 1
        wb.close()

    wb_final.save(ruta_final)
    print(f"Consolidado interno: {ruta_final}")
    return ruta_final


def _consolidar_partes(rutas, encabezado, prefijo):
    """Une partes temporales en un solo libro con varias hojas Datos_N."""
    from openpyxl import load_workbook

    ruta_final = _ruta_temporal(f"{prefijo}cons_")
    wb_final = Workbook(write_only=True)

    for i, ruta in enumerate(rutas, start=1):
        ws = wb_final.create_sheet(f"Datos_{i}")
        wb = load_workbook(ruta, read_only=True, data_only=True)
        hoja = wb.active
        for fila in hoja.iter_rows(values_only=True):
            ws.append(list(fila))
        wb.close()

    wb_final.save(ruta_final)

    for ruta in rutas:
        try:
            os.remove(ruta)
        except OSError:
            pass

    return ruta_final
