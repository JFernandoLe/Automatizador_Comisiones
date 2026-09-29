import os
import re
import calendar
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ============================================================
# CONFIGURACION
# ============================================================

MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12
}

MESES_INV = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
}

errores = []


# ============================================================
# DETECCION PERIODO
# ============================================================

def detectar_periodo_archivo(ruta):

    nombre = os.path.basename(ruta).lower()

    patron = re.search(
        r"(20\d{2})[\s_-]?(0?[1-9]|1[0-2])",
        nombre
    )

    if patron:
        return int(patron.group(1)), int(patron.group(2))

    for texto_mes, numero_mes in MESES.items():

        if texto_mes in nombre:

            anio = re.search(r"(20\d{2})", nombre)

            if anio:
                return int(anio.group(1)), numero_mes

    return None, None


# ============================================================
# PEDIR PERIODO
# ============================================================

def solicitar_periodo():

    ventana = tk.Toplevel(root)
    ventana.title("Periodo del archivo")
    ventana.geometry("300x180")

    resultado = {"anio": None, "mes": None}

    ttk.Label(
        ventana,
        text="No se detectó el periodo"
    ).pack(pady=10)

    combo_mes = ttk.Combobox(
        ventana,
        values=list(MESES.keys()),
        state="readonly"
    )

    combo_mes.pack(pady=5)
    combo_mes.current(0)

    entry_anio = ttk.Entry(ventana)
    entry_anio.pack(pady=5)

    entry_anio.insert(
        0,
        str(datetime.now().year)
    )

    def aceptar():

        resultado["mes"] = MESES[
            combo_mes.get().lower()
        ]

        resultado["anio"] = int(
            entry_anio.get()
        )

        ventana.destroy()

    ttk.Button(
        ventana,
        text="Aceptar",
        command=aceptar
    ).pack(pady=10)

    ventana.grab_set()
    ventana.wait_window()

    return resultado["anio"], resultado["mes"]


# ============================================================
# AJUSTE DE FECHAS
# ============================================================

def ajustar_fecha_pago(fecha, anio, mes):

    if pd.isna(fecha):
        return fecha

    try:

        fecha = pd.to_datetime(fecha)

        inicio_mes = datetime(
            anio,
            mes,
            1
        )

        ultimo = calendar.monthrange(
            anio,
            mes
        )[1]

        fin_mes = datetime(
            anio,
            mes,
            ultimo
        )

        if fecha.year < anio:
            return inicio_mes

        if fecha.year > anio:
            return fin_mes

        if fecha.month < mes:
            return inicio_mes

        if fecha.month > mes:
            return fin_mes

        return fecha

    except:

        return fecha


# ============================================================
# SELECCION DE HOJA
# ============================================================

def seleccionar_hoja(hojas):

    ventana = tk.Toplevel(root)

    ventana.title(
        "Seleccionar hoja"
    )

    ventana.geometry(
        "450x170"
    )

    resultado = {
        "hoja": None
    }

    ttk.Label(
        ventana,
        text="No se encontró la hoja 'Detalle pagado'"
    ).pack(pady=10)

    combo = ttk.Combobox(
        ventana,
        values=hojas,
        width=50,
        state="readonly"
    )

    combo.pack()

    combo.current(0)

    def aceptar():

        resultado["hoja"] = combo.get()
        ventana.destroy()

    ttk.Button(
        ventana,
        text="Aceptar",
        command=aceptar
    ).pack(pady=10)

    ventana.grab_set()
    ventana.wait_window()

    return resultado["hoja"]


# ============================================================
# BUSCAR COLUMNA FECHA
# ============================================================

def encontrar_columna_pago(df):

    posibles = [
        "f. pago",
        "f pago",
        "fecha pago",
        "fecha de pago",
        "pago"
    ]

    for col in df.columns:

        nombre = str(col).lower().strip()

        for p in posibles:

            if p in nombre:
                return col

    return None


# ============================================================
# LIMPIAR RESUMENES Y SUBTOTALES
# ============================================================

def limpiar_filas_resumen(df):

    mask = df.astype(str).apply(
        lambda x: x.str.strip().str.upper()
    )

    subtotal_rows = mask.eq(
        "SUBTOTAL"
    ).any(axis=1)

    if subtotal_rows.any():

        idx = subtotal_rows.idxmax()

        return df.loc[:idx - 1]

    return df


# ============================================================
# LEER EXCEL
# ============================================================

def leer_excel(ruta):

    try:

        excel = pd.ExcelFile(ruta)

        hoja = None

        for h in excel.sheet_names:

            if h.strip().lower() == "detalle pagado":

                hoja = h
                break

        if hoja is None:

            hoja = seleccionar_hoja(
                excel.sheet_names
            )

        if hoja is None:
            return None

        df = pd.read_excel(
            ruta,
            sheet_name=hoja,
            header=8
        )

        df.columns = [
            str(c).strip()
            for c in df.columns
        ]

        # eliminar filas completamente vacías
        df = df.dropna(how="all")

        # eliminar filas donde TODAS las columnas estén vacías
        df = df[
            df.notna().any(axis=1)
        ]

        # cortar en SUBTOTAL / KPI / REGION
        df = limpiar_filas_resumen(df)

        # resetear indices
        df.reset_index(
            drop=True,
            inplace=True
        )

        return df

    except Exception as e:

        errores.append({
            "Archivo": ruta,
            "Error": str(e)
        })

        return None


# ============================================================
# ARCHIVOS
# ============================================================

def obtener_archivos(carpeta):

    salida = []

    extensiones = (
        ".xlsx",
        ".xls",
        ".xlsm"
    )

    for raiz, _, archivos in os.walk(carpeta):

        for archivo in archivos:

            if archivo.lower().endswith(
                extensiones
            ):

                salida.append(
                    os.path.join(
                        raiz,
                        archivo
                    )
                )

    return salida


# ============================================================
# PROCESAR
# ============================================================

def procesar_archivo(
    ruta,
    carpeta_csv,
    acumulado_csv
):

    df = leer_excel(ruta)

    if df is None:
        return False

    anio, mes = detectar_periodo_archivo(
        ruta
    )

    if anio is None:

        anio, mes = solicitar_periodo()

    col_pago = encontrar_columna_pago(
        df
    )

    if col_pago is None:

        errores.append({
            "Archivo": ruta,
            "Error": "No se localizó columna F. Pago"
        })

        return False

    df[col_pago] = pd.to_datetime(
        df[col_pago],
        errors="coerce"
    )

    df[col_pago] = df[col_pago].apply(
        lambda x:
        ajustar_fecha_pago(
            x,
            anio,
            mes
        )
    )

    df["Archivo_Origen"] = os.path.basename(
        ruta
    )

    df["Ruta_Origen"] = ruta

    df["Año_Proceso"] = anio

    df["Mes_Proceso"] = mes

    df["Periodo"] = f"{anio}-{mes:02d}"

    df["Fecha_Carga"] = datetime.now()

    nombre = os.path.splitext(
        os.path.basename(ruta)
    )[0]

    csv_individual = os.path.join(
        carpeta_csv,
        f"{nombre}.csv"
    )

    df.to_csv(
        csv_individual,
        index=False,
        encoding="utf-8-sig"
    )

    existe = os.path.exists(
        acumulado_csv
    )

    df.to_csv(
        acumulado_csv,
        mode="a",
        header=not existe,
        index=False,
        encoding="utf-8-sig"
    )

    return True


# ============================================================
# EJECUTAR
# ============================================================

def ejecutar():

    try:

        if modo.get() == "Excel":

            archivo = filedialog.askopenfilename(
                filetypes=[
                    (
                        "Excel",
                        "*.xlsx *.xls *.xlsm"
                    )
                ]
            )

            if not archivo:
                return

            archivos = [archivo]

        else:

            carpeta = filedialog.askdirectory()

            if not carpeta:
                return

            archivos = obtener_archivos(
                carpeta
            )

            if not archivos:

                messagebox.showwarning(
                    "Aviso",
                    "No se encontraron Excel."
                )

                return

        salida = filedialog.askdirectory(
            title="Carpeta de salida"
        )

        if not salida:
            return

        carpeta_csv = os.path.join(
            salida,
            "CSV_Individuales"
        )

        os.makedirs(
            carpeta_csv,
            exist_ok=True
        )

        acumulado_csv = os.path.join(
            salida,
            "ACUMULADO_FINAL.csv"
        )

        if os.path.exists(
            acumulado_csv
        ):

            os.remove(
                acumulado_csv
            )

        barra["maximum"] = len(
            archivos
        )

        procesados = 0

        for i, archivo in enumerate(
            archivos,
            start=1
        ):

            estado.set(
                f"Procesando {i} de {len(archivos)}"
            )

            root.update_idletasks()

            ok = procesar_archivo(
                archivo,
                carpeta_csv,
                acumulado_csv
            )

            if ok:
                procesados += 1

            barra["value"] = i

        if errores:

            pd.DataFrame(
                errores
            ).to_excel(
                os.path.join(
                    salida,
                    "LOG_ERRORES.xlsx"
                ),
                index=False
            )

        estado.set(
            "Proceso terminado"
        )

        messagebox.showinfo(
            "Finalizado",
            f"Archivos procesados: {procesados}"
        )

    except Exception as ex:

        messagebox.showerror(
            "Error",
            str(ex)
        )


# ============================================================
# GUI
# ============================================================

root = tk.Tk()

root.title(
    "Extractor Detalle Pagado"
)

root.geometry(
    "650x320"
)

ttk.Label(
    root,
    text="Tipo de proceso",
    font=("Segoe UI", 12, "bold")
).pack(
    pady=15
)

modo = tk.StringVar(
    value="Carpeta"
)

ttk.Radiobutton(
    root,
    text="Procesar carpeta",
    variable=modo,
    value="Carpeta"
).pack()

ttk.Radiobutton(
    root,
    text="Procesar un Excel",
    variable=modo,
    value="Excel"
).pack()

ttk.Button(
    root,
    text="Ejecutar",
    command=ejecutar
).pack(
    pady=20
)

barra = ttk.Progressbar(
    root,
    length=550
)

barra.pack(
    pady=10
)

estado = tk.StringVar(
    value="Esperando..."
)

ttk.Label(
    root,
    textvariable=estado
).pack()

root.mainloop()