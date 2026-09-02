import tkinter as tk
from tkinter import ttk


def crear_area_desplazable(parent):
    canvas = tk.Canvas(parent, bg="#F5F7FA", highlightthickness=0)
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    contenido = ttk.Frame(canvas)

    ventana = canvas.create_window((0, 0), window=contenido, anchor="nw")
    contenido.bind(
        "<Configure>",
        lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.bind(
        "<Configure>",
        lambda e: canvas.itemconfigure(ventana, width=e.width),
    )
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def desplazamiento(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", desplazamiento)
    return contenido


def crear_selector_archivo(parent, titulo, comando, ancho=80):
    ttk.Label(parent, text=titulo).pack(pady=(10, 5))
    frame = ttk.Frame(parent)
    frame.pack(fill="x", padx=20)
    entrada = ttk.Entry(frame, width=ancho)
    entrada.pack(side="left", expand=True, fill="x")
    ttk.Button(frame, text="Buscar", command=comando).pack(side="left", padx=5)
    return entrada


def crear_selector_multiples(parent, titulo, comando_archivos, comando_carpeta, ancho=80):
    ttk.Label(parent, text=titulo).pack(pady=(10, 5))
    frame = ttk.Frame(parent)
    frame.pack(fill="x", padx=20)
    entrada = ttk.Entry(frame, width=ancho)
    entrada.pack(side="left", expand=True, fill="x")
    ttk.Button(frame, text="Archivos", command=comando_archivos).pack(side="left", padx=5)
    ttk.Button(frame, text="Carpeta", command=comando_carpeta).pack(side="left", padx=5)
    return entrada


def reemplazar_texto(entrada, texto):
    entrada.delete(0, tk.END)
    entrada.insert(0, texto)


def llenar_checks(parent, hojas):
    for widget in parent.winfo_children():
        widget.destroy()

    checks = []
    for hoja in hojas:
        variable = tk.BooleanVar(value=True)
        tk.Checkbutton(parent, text=hoja, variable=variable).pack(anchor="w")
        checks.append((hoja, variable))
    return checks
