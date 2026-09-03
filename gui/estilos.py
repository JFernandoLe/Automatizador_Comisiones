COLORES = {
    "header": "#000000",
    "fondo": "#F3F5F7",
    "tarjeta": "#FFFFFF",
    "texto": "#1A1F24",
    "muted": "#5C6770",
    "primario": "#009CDE",
    "primario_hover": "#0086C2",
    "acento": "#B5D334",
    "borde": "#E2E7EC",
    "exito": "#1B7A4E",
    "error": "#C0392B",
    "pista": "#E6EAEF",
}


def aplicar_estilos(root, style):
    root.configure(bg=COLORES["fondo"])
    style.theme_use("clam")

    style.configure(".", font=("Segoe UI", 10), background=COLORES["fondo"])
    style.configure("TFrame", background=COLORES["fondo"])
    style.configure("Fondo.TFrame", background=COLORES["fondo"])
    style.configure("Card.TFrame", background=COLORES["tarjeta"])
    style.configure(
        "TLabel",
        background=COLORES["fondo"],
        foreground=COLORES["texto"],
        font=("Segoe UI", 10),
    )
    style.configure(
        "Card.TLabel",
        background=COLORES["tarjeta"],
        foreground=COLORES["texto"],
        font=("Segoe UI", 10),
    )
    style.configure(
        "CardTitle.TLabel",
        background=COLORES["tarjeta"],
        foreground=COLORES["texto"],
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "Muted.TLabel",
        background=COLORES["tarjeta"],
        foreground=COLORES["muted"],
        font=("Segoe UI", 9),
    )
    style.configure(
        "Hint.TLabel",
        background=COLORES["fondo"],
        foreground=COLORES["muted"],
        font=("Segoe UI", 9),
    )
    style.configure(
        "Section.TLabel",
        background=COLORES["fondo"],
        foreground=COLORES["texto"],
        font=("Segoe UI Semibold", 12),
    )
    style.configure(
        "Status.TLabel",
        background=COLORES["tarjeta"],
        foreground=COLORES["texto"],
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "Percent.TLabel",
        background=COLORES["tarjeta"],
        foreground=COLORES["primario"],
        font=("Segoe UI Semibold", 12),
    )
    style.configure(
        "TLabelframe",
        background=COLORES["tarjeta"],
        foreground=COLORES["texto"],
        bordercolor=COLORES["borde"],
        relief="flat",
        padding=8,
    )
    style.configure(
        "TLabelframe.Label",
        background=COLORES["tarjeta"],
        foreground=COLORES["muted"],
        font=("Segoe UI Semibold", 9),
    )
    style.configure(
        "Card.TLabelframe",
        background=COLORES["tarjeta"],
        bordercolor=COLORES["borde"],
        relief="solid",
        padding=10,
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=COLORES["tarjeta"],
        foreground=COLORES["muted"],
        font=("Segoe UI Semibold", 9),
    )
    style.configure(
        "TButton",
        font=("Segoe UI Semibold", 9),
        padding=(12, 7),
        background="#EEF2F5",
        foreground=COLORES["texto"],
        bordercolor=COLORES["borde"],
        relief="flat",
    )
    style.map(
        "TButton",
        background=[("active", "#E2E8ED"), ("disabled", "#F3F5F7")],
        foreground=[("disabled", "#9AA3AB")],
    )
    style.configure(
        "TEntry",
        fieldbackground="#FAFBFC",
        bordercolor=COLORES["borde"],
        lightcolor=COLORES["borde"],
        darkcolor=COLORES["borde"],
        padding=6,
        font=("Segoe UI", 9),
    )
    style.configure(
        "TCheckbutton",
        background=COLORES["tarjeta"],
        foreground=COLORES["texto"],
        font=("Segoe UI", 9),
        padding=2,
    )
    style.map(
        "TCheckbutton",
        background=[("active", COLORES["tarjeta"])],
    )
    style.configure(
        "Commi.Horizontal.TProgressbar",
        troughcolor=COLORES["pista"],
        background=COLORES["primario"],
        lightcolor=COLORES["primario"],
        darkcolor=COLORES["primario"],
        bordercolor=COLORES["pista"],
        thickness=16,
    )
    style.configure("TScrollbar", background=COLORES["fondo"], troughcolor=COLORES["fondo"])
    style.configure(
        "TNotebook",
        background=COLORES["fondo"],
        borderwidth=0,
        tabmargins=(12, 8, 12, 0),
    )
    style.configure(
        "TNotebook.Tab",
        background="#E4EAEE",
        foreground=COLORES["muted"],
        padding=(18, 8),
        font=("Segoe UI Semibold", 10),
        borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", COLORES["tarjeta"]), ("active", "#EDF2F5")],
        foreground=[("selected", COLORES["primario"]), ("active", COLORES["texto"])],
    )
