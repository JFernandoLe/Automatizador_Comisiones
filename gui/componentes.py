import tkinter as tk
from tkinter import ttk

from gui.estilos import COLORES


def crear_area_desplazable(parent, usar_rueda=False):
    canvas = tk.Canvas(parent, bg=COLORES["fondo"], highlightthickness=0, bd=0)
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    contenido = ttk.Frame(canvas, style="Fondo.TFrame")

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

    def _puntero_sobre_el_area():
        if not parent.winfo_ismapped():
            return False
        x, y = parent.winfo_pointerxy()
        izquierda = parent.winfo_rootx()
        arriba = parent.winfo_rooty()
        return (
            izquierda <= x <= izquierda + parent.winfo_width()
            and arriba <= y <= arriba + parent.winfo_height()
        )

    def desplazamiento(event):
        if not _puntero_sobre_el_area():
            return
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    if usar_rueda:
        canvas.bind_all("<MouseWheel>", desplazamiento)

    return contenido


def _tarjeta(parent):
    exterior = tk.Frame(parent, bg=COLORES["fondo"])
    exterior.pack(fill="x", padx=28, pady=6)
    borde = tk.Frame(exterior, bg=COLORES["borde"])
    borde.pack(fill="x")
    tarjeta = tk.Frame(borde, bg=COLORES["tarjeta"])
    tarjeta.pack(fill="x", padx=1, pady=1)
    return tarjeta


def crear_selector_archivo(parent, titulo, comando, ancho=80, ayuda=None):
    tarjeta = _tarjeta(parent)
    ttk.Label(tarjeta, text=titulo, style="CardTitle.TLabel").pack(
        anchor="w", padx=16, pady=(12, 2)
    )
    if ayuda:
        ttk.Label(tarjeta, text=ayuda, style="Muted.TLabel").pack(anchor="w", padx=16)
    frame = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
    frame.pack(fill="x", padx=16, pady=(8, 14))
    entrada = ttk.Entry(frame, width=ancho)
    entrada.pack(side="left", expand=True, fill="x")
    ttk.Button(frame, text="Seleccionar", command=comando).pack(side="left", padx=(8, 0))
    return entrada


def crear_selector_multiples(
    parent, titulo, comando_archivos, comando_carpeta, ancho=80, ayuda=None
):
    tarjeta = _tarjeta(parent)
    ttk.Label(tarjeta, text=titulo, style="CardTitle.TLabel").pack(
        anchor="w", padx=16, pady=(12, 2)
    )
    if ayuda:
        ttk.Label(tarjeta, text=ayuda, style="Muted.TLabel").pack(anchor="w", padx=16)
    frame = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
    frame.pack(fill="x", padx=16, pady=(8, 14))
    entrada = ttk.Entry(frame, width=ancho)
    entrada.pack(side="left", expand=True, fill="x")
    ttk.Button(frame, text="Archivos", command=comando_archivos).pack(
        side="left", padx=(8, 0)
    )
    ttk.Button(frame, text="Carpeta", command=comando_carpeta).pack(
        side="left", padx=(8, 0)
    )
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
        ttk.Checkbutton(parent, text=hoja, variable=variable).pack(anchor="w", padx=4)
        checks.append((hoja, variable))
    return checks
