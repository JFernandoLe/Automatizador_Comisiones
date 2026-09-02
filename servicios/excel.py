from pathlib import Path

import pandas as pd
from python_calamine import CalamineWorkbook

EXCEL_MAX_ROWS = 1_048_000


def _es_xls(archivo):
    return Path(archivo).suffix.lower() == ".xls"


def _engine(archivo):
    if _es_xls(archivo):
        return "calamine"
    return "openpyxl"


def obtener_hojas(archivo):
    return list(CalamineWorkbook.from_path(str(archivo)).sheet_names)


def listar_excel_en_carpeta(carpeta):
    raiz = Path(carpeta)
    archivos = []
    for patron in ("*.xls", "*.xlsx"):
        archivos.extend(raiz.glob(patron))
    archivos = [
        str(a.resolve())
        for a in sorted(archivos)
        if not a.name.startswith("~$")
        and not a.name.lower().startswith("consolidado")
    ]
    return archivos


def obtener_hojas_union(archivos):
    hojas = []
    vistas = set()
    for archivo in archivos:
        for hoja in obtener_hojas(archivo):
            if hoja not in vistas:
                vistas.add(hoja)
                hojas.append(hoja)
    return hojas


def leer_hojas_seleccionadas(archivo, hojas_seleccionadas, header=0):
    if not hojas_seleccionadas:
        raise ValueError("Debe seleccionar al menos una hoja.")

    dataframes = []
    engine = _engine(archivo)
    for hoja in hojas_seleccionadas:
        print("\n" + "=" * 80)
        print(f"Leyendo hoja: {hoja}")
        df = pd.read_excel(
            archivo,
            sheet_name=hoja,
            engine=engine,
            header=header,
        )
        print(f"Registros hoja: {len(df):,}")
        dataframes.append(df)

    resultado = pd.concat(dataframes, ignore_index=True)
    print(f"Total consolidado: {len(resultado):,}")
    return resultado


def guardar_dataframe_temporal(df, prefijo="tmp_"):
    import os
    import tempfile

    fd, ruta = tempfile.mkstemp(prefix=prefijo, suffix=".xlsx")
    os.close(fd)
    df.to_excel(ruta, index=False, engine="openpyxl")
    return ruta


def guardar_excel_dividido(df, archivo):
    total_filas = len(df)
    hojas_necesarias = max(1, (total_filas + EXCEL_MAX_ROWS - 1) // EXCEL_MAX_ROWS)

    print("\n" + "=" * 80)
    print(f"GENERANDO EXCEL: {archivo}")
    print(f"Total registros: {total_filas:,}")
    print(f"Hojas requeridas: {hojas_necesarias}")

    with pd.ExcelWriter(
        archivo,
        engine="xlsxwriter",
        datetime_format="dd/mm/yyyy",
    ) as writer:
        workbook = writer.book
        formato_fecha = workbook.add_format({"num_format": "dd/mm/yyyy"})

        for i in range(hojas_necesarias):
            inicio = i * EXCEL_MAX_ROWS
            fin = min((i + 1) * EXCEL_MAX_ROWS, total_filas)
            df_export = df.iloc[inicio:fin].copy()

            if "FECHA" in df_export.columns:
                df_export["FECHA"] = pd.to_datetime(
                    df_export["FECHA"], errors="coerce"
                )

            nombre_hoja = f"Datos_{i + 1}"
            df_export.to_excel(writer, sheet_name=nombre_hoja, index=False)
            worksheet = writer.sheets[nombre_hoja]

            if "FECHA" in df_export.columns:
                col_fecha = df_export.columns.get_loc("FECHA")
                worksheet.set_column(col_fecha, col_fecha, 15, formato_fecha)

    print(f"Excel generado correctamente: {archivo}")
