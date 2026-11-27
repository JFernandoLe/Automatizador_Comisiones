import tkinter as tk
from tkinter import ttk

from gui.estilos import COLORES

ETIQUETAS = {
    "requerido": ("Requerido", COLORES["primario"], COLORES["info_fondo"]),
    "opcional": ("Opcional", COLORES["muted"], COLORES["neutro_fondo"]),
    "fy": ("Año completo", COLORES["aviso"], COLORES["aviso_fondo"]),
    "auto": ("Se llena solo", COLORES["exito"], COLORES["exito_fondo"]),
}


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


def crear_etiqueta(parent, tipo, fondo=None):
    texto, color, color_fondo = ETIQUETAS[tipo]
    return tk.Label(
        parent,
        text=texto,
        fg=color,
        bg=color_fondo,
        font=("Segoe UI Semibold", 8),
        padx=8,
        pady=1,
    )


def _titulo_con_etiquetas(tarjeta, titulo, etiquetas):
    fila = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
    fila.pack(fill="x", padx=16, pady=(12, 2))
    ttk.Label(fila, text=titulo, style="CardTitle.TLabel").pack(side="left")
    for tipo in etiquetas or ():
        crear_etiqueta(fila, tipo).pack(side="left", padx=(8, 0))


def crear_selector_archivo(
    parent, titulo, comando, ancho=80, ayuda=None, etiquetas=None
):
    tarjeta = _tarjeta(parent)
    _titulo_con_etiquetas(tarjeta, titulo, etiquetas)
    if ayuda:
        ttk.Label(
            tarjeta, text=ayuda, style="Muted.TLabel", wraplength=860, justify="left"
        ).pack(anchor="w", padx=16)
    frame = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
    frame.pack(fill="x", padx=16, pady=(8, 14))
    entrada = ttk.Entry(frame, width=ancho)
    entrada.pack(side="left", expand=True, fill="x")
    ttk.Button(frame, text="Seleccionar", command=comando).pack(side="left", padx=(8, 0))
    return entrada


def crear_selector_multiples(
    parent,
    titulo,
    comando_archivos,
    comando_carpeta,
    ancho=80,
    ayuda=None,
    etiquetas=None,
):
    tarjeta = _tarjeta(parent)
    _titulo_con_etiquetas(tarjeta, titulo, etiquetas)
    if ayuda:
        ttk.Label(
            tarjeta, text=ayuda, style="Muted.TLabel", wraplength=860, justify="left"
        ).pack(anchor="w", padx=16)
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


def crear_encabezado_paso(
    parent, numero, titulo, descripcion, resultado=None, pasos=None, nota=None
):
    exterior = tk.Frame(parent, bg=COLORES["fondo"])
    exterior.pack(fill="x", padx=28, pady=(18, 10))
    borde = tk.Frame(exterior, bg=COLORES["borde"])
    borde.pack(fill="x")
    tarjeta = tk.Frame(borde, bg=COLORES["tarjeta"])
    tarjeta.pack(fill="x", padx=1, pady=1)
    tk.Frame(tarjeta, bg=COLORES["primario"], width=4).pack(side="left", fill="y")
    cuerpo = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
    cuerpo.pack(fill="x", padx=18, pady=16)

    fila = tk.Frame(cuerpo, bg=COLORES["tarjeta"])
    fila.pack(fill="x")
    if numero:
        tk.Label(
            fila,
            text=str(numero),
            fg="#FFFFFF",
            bg=COLORES["primario"],
            font=("Segoe UI", 12, "bold"),
            width=2,
            pady=2,
        ).pack(side="left", padx=(0, 10))
    tk.Label(
        fila,
        text=titulo,
        fg=COLORES["texto"],
        bg=COLORES["tarjeta"],
        font=("Segoe UI", 15, "bold"),
    ).pack(side="left")

    ttk.Label(
        cuerpo,
        text=descripcion,
        style="Muted.TLabel",
        wraplength=880,
        justify="left",
    ).pack(anchor="w", pady=(8, 0))

    if resultado:
        fila_res = tk.Frame(cuerpo, bg=COLORES["tarjeta"])
        fila_res.pack(anchor="w", pady=(8, 0))
        tk.Label(
            fila_res,
            text="Resultado:",
            fg=COLORES["texto"],
            bg=COLORES["tarjeta"],
            font=("Segoe UI Semibold", 9),
        ).pack(side="left")
        tk.Label(
            fila_res,
            text=resultado,
            fg=COLORES["texto"],
            bg=COLORES["tarjeta"],
            font=("Segoe UI", 9),
        ).pack(side="left", padx=(6, 0))

    if pasos:
        lista = tk.Frame(cuerpo, bg=COLORES["tarjeta"])
        lista.pack(fill="x", pady=(10, 0))
        for indice, texto in enumerate(pasos, start=1):
            renglon = tk.Frame(lista, bg=COLORES["tarjeta"])
            renglon.pack(fill="x", pady=1)
            tk.Label(
                renglon,
                text=f"{indice}.",
                fg=COLORES["primario"],
                bg=COLORES["tarjeta"],
                font=("Segoe UI Semibold", 9),
                width=3,
                anchor="ne",
            ).pack(side="left", anchor="n")
            tk.Label(
                renglon,
                text=texto,
                fg=COLORES["texto"],
                bg=COLORES["tarjeta"],
                font=("Segoe UI", 9),
                wraplength=840,
                justify="left",
                anchor="w",
            ).pack(side="left", fill="x", anchor="w")

    if nota:
        crear_nota(cuerpo, nota, dentro_de_tarjeta=True)
    return cuerpo


def crear_nota(parent, texto, tipo="aviso", dentro_de_tarjeta=False):
    colores = {
        "aviso": (COLORES["aviso"], COLORES["aviso_fondo"]),
        "info": (COLORES["primario"], COLORES["info_fondo"]),
        "exito": (COLORES["exito"], COLORES["exito_fondo"]),
    }
    color, fondo = colores.get(tipo, colores["aviso"])
    caja = tk.Frame(parent, bg=fondo)
    if dentro_de_tarjeta:
        caja.pack(fill="x", pady=(12, 0))
    else:
        caja.pack(fill="x", padx=28, pady=(4, 8))
    tk.Frame(caja, bg=color, width=3).pack(side="left", fill="y")
    tk.Label(
        caja,
        text=texto,
        fg=COLORES["texto"],
        bg=fondo,
        font=("Segoe UI", 9),
        wraplength=860,
        justify="left",
        anchor="w",
        padx=10,
        pady=8,
    ).pack(side="left", fill="x", expand=True)
    return caja


def crear_seccion(parent, titulo, descripcion=None, primera=False):
    contenedor = tk.Frame(parent, bg=COLORES["fondo"])
    contenedor.pack(fill="x", padx=28, pady=((8 if primera else 22), 2))
    ttk.Label(contenedor, text=titulo, style="Section.TLabel").pack(anchor="w")
    if descripcion:
        ttk.Label(
            contenedor,
            text=descripcion,
            style="Hint.TLabel",
            wraplength=880,
            justify="left",
        ).pack(anchor="w", pady=(2, 4))
    tk.Frame(contenedor, bg=COLORES["borde"], height=1).pack(fill="x", pady=(4, 0))
    return contenedor


def crear_boton_primario(parent, texto, comando, ayuda=None):
    marco = tk.Frame(parent, bg=COLORES["fondo"])
    marco.pack(pady=(26, 38))
    boton = tk.Button(
        marco,
        text=texto,
        command=comando,
        bg=COLORES["primario"],
        fg="#FFFFFF",
        activebackground=COLORES["primario_hover"],
        activeforeground="#FFFFFF",
        font=("Segoe UI Semibold", 12),
        relief="flat",
        bd=0,
        cursor="hand2",
        padx=36,
        pady=12,
    )
    boton.pack()
    boton.bind(
        "<Enter>",
        lambda _e: boton.config(bg=COLORES["primario_hover"])
        if str(boton["state"]) == "normal"
        else None,
    )
    boton.bind(
        "<Leave>",
        lambda _e: boton.config(bg=COLORES["primario"])
        if str(boton["state"]) == "normal"
        else None,
    )
    if ayuda:
        ttk.Label(marco, text=ayuda, style="Hint.TLabel").pack(pady=(8, 0))
    return boton


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
