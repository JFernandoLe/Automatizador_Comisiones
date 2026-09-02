import pandas as pd

EXCEL_MAX_ROWS = 1_048_000


def obtener_hojas(archivo):
    return pd.ExcelFile(archivo).sheet_names


def leer_hojas_seleccionadas(archivo, hojas_seleccionadas):
    if not hojas_seleccionadas:
        raise ValueError("Debe seleccionar al menos una hoja.")

    dataframes = []
    for hoja in hojas_seleccionadas:
        print("\n" + "=" * 80)
        print(f"Leyendo hoja: {hoja}")
        df = pd.read_excel(archivo, sheet_name=hoja, engine="openpyxl")
        print(f"Registros hoja: {len(df):,}")
        dataframes.append(df)

    resultado = pd.concat(dataframes, ignore_index=True)
    print(f"Total consolidado: {len(resultado):,}")
    return resultado


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
