from pathlib import Path

from ingresos.excel_ingresos import llenar_plantilla
from ingresos.fuentes import consolidar_datos

RUTA_PLANTILLA = (
    Path(__file__).resolve().parents[1] / "plantillas" / "Ingresos_vs_Ventas.xlsx"
)


def generar_ingresos(
    anio_actual,
    mes_corte,
    archivo_actual,
    archivo_anterior=None,
    archivo_dos_anios=None,
    ruta_salida="Ingresos vs Ventas.xlsx",
    ruta_plantilla=None,
    actualizar_estado=None,
):
    if not anio_actual:
        raise ValueError("Debe indicar el año actual del reporte.")
    mes_corte = int(mes_corte)
    if mes_corte < 1 or mes_corte > 12:
        raise ValueError("El mes de corte debe estar entre 1 y 12.")
    plantilla = Path(ruta_plantilla) if ruta_plantilla else RUTA_PLANTILLA
    if not plantilla.exists():
        raise ValueError(f"No se encontró la plantilla: {plantilla}")
    if not archivo_actual:
        raise ValueError("Debe seleccionar el Excel del año actual (CommiFlow o bases).")

    def avisar(texto, progreso=None):
        if actualizar_estado:
            actualizar_estado(texto, progreso)

    avisar("Consolidando bases de Ingresos vs Ventas...", 8)
    paquetes = [
        (int(anio_actual), archivo_actual),
        (int(anio_actual) - 1, archivo_anterior),
        (int(anio_actual) - 2, archivo_dos_anios),
    ]
    datos = consolidar_datos(paquetes, actualizar_estado=actualizar_estado)
    avisar("Llenando plantilla...", 72)
    ruta = llenar_plantilla(
        plantilla,
        ruta_salida,
        datos,
        int(anio_actual),
        mes_corte,
    )
    avisar("Ingresos vs Ventas generado", 100)
    return ruta
