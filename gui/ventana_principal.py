import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from conversores.preparar import es_entrada_txt
from gui.componentes import (
    crear_area_desplazable,
    crear_selector_archivo,
    crear_selector_multiples,
    llenar_checks,
    reemplazar_texto,
)
from gui.estilos import COLORES, aplicar_estilos
from procesos.bases_sap import generar_bases_sap
from procesos.comisiones import generar_comisiones
from procesos.reporte_final import generar_reporte_final
from servicios.excel import listar_excel_en_carpeta, obtener_hojas, obtener_hojas_union
from servicios.recursos import ruta_logo

TIPOS_EXCEL = [("Excel", "*.xlsx *.xls")]
TIPOS_SAA = [
    ("SAA", "*.txt *.xlsx *.xls"),
    ("Texto", "*.txt"),
    ("Excel", "*.xlsx *.xls"),
]
CLAVES_MULTIPLES = ("vida", "gmm", "saa", "manuales")


class VentanaPrincipal:
    def __init__(self, root):
        self.root = root
        self.archivos = {
            "vida": None,
            "gmm": None,
            "saa": None,
            "manuales": None,
            "vlsp": None,
            "tipo": None,
            "catalogos": None,
        }
        self.checks_vida = []
        self.checks_gmm = []
        self.checks_saa = []
        self.checks_manuales = []
        self.checks_vlsp = []
        self.checks_tipo = []
        self.checks_catalogos = []

        self.checks_pc_vida = []
        self.checks_pc_gmm = []
        self.checks_pc_saa = []
        self.checks_pc_manuales = []
        self.checks_pc_vlsp = []
        self.checks_pc_tipo = []
        self.checks_pc_catalogos = []

        self.modo_proceso_completo = False
        self._progreso_global = False
        self._progreso_minimo = 0
        self._logo_header = None
        self._logo_acerca = None

        self.generar_vida_var = tk.BooleanVar(value=True)
        self.generar_gmm_var = tk.BooleanVar(value=True)
        self.progress_var = tk.IntVar(value=0)

        self._configurar_ventana()
        self._crear_menu()
        self._crear_interfaz()

    def _configurar_ventana(self):
        self.root.title("CommiFlow")
        self.root.geometry("1080x780")
        self.root.minsize(920, 680)
        style = ttk.Style()
        aplicar_estilos(self.root, style)

    def _crear_menu(self):
        menubar = tk.Menu(self.root)
        menu_ayuda = tk.Menu(menubar, tearoff=0)
        menu_ayuda.add_command(
            label="Acerca de CommiFlow",
            command=self._mostrar_acerca_de,
        )
        menubar.add_cascade(label="Ayuda", menu=menu_ayuda)
        self.root.config(menu=menubar)

    def _mostrar_acerca_de(self):
        self.notebook.select(self.tab_acerca)

    def _crear_interfaz(self):
        self._crear_header()
        self._crear_footer()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        self.tab_completo = ttk.Frame(self.notebook, style="Fondo.TFrame")
        self.tab_acerca = ttk.Frame(self.notebook, style="Fondo.TFrame")
        self.notebook.add(self.tab_completo, text="Proceso")
        self.notebook.add(self.tab_acerca, text="Acerca de")

        self.tab_bases = ttk.Frame(self.root)
        self.tab_comisiones = ttk.Frame(self.root)
        self.tab_reporte = ttk.Frame(self.root)

        self._crear_tab_completo()
        self._crear_tab_acerca()
        self._crear_tab_bases()
        self._crear_tab_comisiones()
        self._crear_tab_reporte()

    def _crear_header(self):
        header = tk.Frame(self.root, bg=COLORES["header"], height=92)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        logo = self._cargar_logo(alto=64, destino="_logo_header")
        if logo is not None:
            tk.Label(header, image=logo, bg=COLORES["header"], bd=0).pack(
                side="left", padx=(22, 18), pady=12
            )

        textos = tk.Frame(header, bg=COLORES["header"])
        textos.pack(side="left", pady=16)
        tk.Label(
            textos,
            text="CommiFlow",
            fg="#FFFFFF",
            bg=COLORES["header"],
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            textos,
            text="Automatización y procesamiento de comisiones",
            fg="#C5CBD1",
            bg=COLORES["header"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(2, 0))

        tk.Button(
            header,
            text="Acerca de",
            command=self._mostrar_acerca_de,
            bg=COLORES["header"],
            fg="#FFFFFF",
            activebackground="#1A1A1A",
            activeforeground="#FFFFFF",
            font=("Segoe UI Semibold", 10),
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=16,
            pady=8,
        ).pack(side="right", padx=22)

    def _cargar_logo(self, alto=64, destino="_logo_header"):
        ruta = ruta_logo()
        if ruta is None:
            return None
        try:
            from PIL import Image, ImageTk

            imagen = Image.open(ruta)
            ancho = max(1, int(imagen.width * (alto / imagen.height)))
            imagen = imagen.resize((ancho, alto), Image.Resampling.LANCZOS)
            foto = ImageTk.PhotoImage(imagen)
            setattr(self, destino, foto)
            return foto
        except Exception:
            return None

    def _crear_tab_completo(self):
        contenido = crear_area_desplazable(self.tab_completo, usar_rueda=True)
        self._crear_instrucciones(contenido)

        ttk.Label(contenido, text="Archivos de entrada", style="Section.TLabel").pack(
            anchor="w", padx=28, pady=(8, 2)
        )
        ttk.Label(
            contenido,
            text="Seleccione cada archivo o carpeta que CommiFlow necesita para el proceso.",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=28, pady=(0, 8))

        self.entrada_pc_vida = crear_selector_multiples(
            contenido,
            "Base VIDA",
            lambda: self._seleccionar_archivos("vida"),
            lambda: self._seleccionar_carpeta("vida"),
            ayuda="Uno o varios Excel originales, o una carpeta.",
        )
        self.frame_pc_vida_hojas = self._marco_hojas(contenido, "Hojas de Base VIDA")

        self.entrada_pc_gmm = crear_selector_multiples(
            contenido,
            "Base GMM",
            lambda: self._seleccionar_archivos("gmm"),
            lambda: self._seleccionar_carpeta("gmm"),
            ayuda="Uno o varios Excel originales, o una carpeta.",
        )
        self.frame_pc_gmm_hojas = self._marco_hojas(contenido, "Hojas de Base GMM")

        self.entrada_pc_saa = crear_selector_multiples(
            contenido,
            "Archivo SAA",
            lambda: self._seleccionar_archivos("saa"),
            lambda: self._seleccionar_carpeta("saa"),
            ayuda="Un archivo TXT, o uno o varios Excel.",
        )
        self.frame_pc_saa_hojas = self._marco_hojas(contenido, "Hojas de SAA")

        self.entrada_pc_manuales = crear_selector_multiples(
            contenido,
            "Acumulado de Comisiones",
            lambda: self._seleccionar_archivos("manuales"),
            lambda: self._seleccionar_carpeta("manuales"),
            ayuda="Uno o varios Excel originales, o una carpeta.",
        )
        self.frame_pc_manuales_hojas = self._marco_hojas(
            contenido, "Hojas de Acumulado de Comisiones"
        )

        ttk.Label(
            contenido, text="Archivos de referencia y catálogos", style="Section.TLabel"
        ).pack(anchor="w", padx=28, pady=(18, 2))
        ttk.Label(
            contenido,
            text="Estos archivos se utilizan al generar el reporte final.",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=28, pady=(0, 8))

        self.entrada_pc_vlsp = crear_selector_archivo(
            contenido,
            "Archivo VLSP",
            lambda: self._seleccionar_archivo("vlsp"),
            ayuda="Seleccione un archivo Excel.",
        )
        self.frame_pc_vlsp_hojas = self._marco_hojas(contenido, "Hojas VLSP")

        self.entrada_pc_tipo = crear_selector_archivo(
            contenido,
            "Catálogo Estatus Pólizas",
            lambda: self._seleccionar_archivo("tipo"),
            ayuda="Seleccione un archivo Excel.",
        )
        self.frame_pc_tipo_hojas = self._marco_hojas(
            contenido, "Hojas de Catálogo Estatus Pólizas"
        )

        self.entrada_pc_catalogos = crear_selector_archivo(
            contenido,
            "Archivo único de Catálogos",
            lambda: self._seleccionar_archivo("catalogos"),
            ayuda="Debe incluir las hojas PFPM, CONCEPTOS_VIDA y CONCEPTOS_GMM.",
        )
        self.frame_pc_catalogos_hojas = self._marco_hojas(contenido, "Hojas de Catálogos")

        boton_frame = tk.Frame(contenido, bg=COLORES["fondo"])
        boton_frame.pack(pady=(24, 36))
        self.boton_proceso_completo = tk.Button(
            boton_frame,
            text="Ejecutar CommiFlow",
            command=self.ejecutar_proceso_completo,
            bg=COLORES["primario"],
            fg="#FFFFFF",
            activebackground=COLORES["primario_hover"],
            activeforeground="#FFFFFF",
            font=("Segoe UI Semibold", 12),
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=32,
            pady=11,
        )
        self.boton_proceso_completo.pack()
        self.boton_proceso_completo.bind(
            "<Enter>",
            lambda _e: self.boton_proceso_completo.config(
                bg=COLORES["primario_hover"]
            )
            if str(self.boton_proceso_completo["state"]) == "normal"
            else None,
        )
        self.boton_proceso_completo.bind(
            "<Leave>",
            lambda _e: self.boton_proceso_completo.config(bg=COLORES["primario"])
            if str(self.boton_proceso_completo["state"]) == "normal"
            else None,
        )

    def _crear_instrucciones(self, parent):
        exterior = tk.Frame(parent, bg=COLORES["fondo"])
        exterior.pack(fill="x", padx=28, pady=(18, 10))
        borde = tk.Frame(exterior, bg=COLORES["borde"])
        borde.pack(fill="x")
        tarjeta = tk.Frame(borde, bg=COLORES["tarjeta"])
        tarjeta.pack(fill="x", padx=1, pady=1)

        acento = tk.Frame(tarjeta, bg=COLORES["primario"], width=4)
        acento.pack(side="left", fill="y")

        cuerpo = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
        cuerpo.pack(fill="x", padx=16, pady=14)

        ttk.Label(
            cuerpo,
            text="¿Cómo utilizar CommiFlow?",
            style="CardTitle.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            cuerpo,
            text="Siga estos pasos en el orden indicado. El proceso no cambia: solo se muestra con mayor claridad.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        pasos = [
            (
                "1. Selecciona los archivos de entrada",
                "Carga Base VIDA, Base GMM, el archivo SAA (TXT o Excel), Acumulado de Comisiones, VLSP, Catálogo Estatus Pólizas y el archivo único de Catálogos.",
            ),
            (
                "2. Selecciona las hojas correspondientes",
                "Si un Excel tiene varias hojas, marque cuáles desea procesar. En carpeta o varios archivos de VIDA, GMM o Acumulado se consolidan todas las hojas de cada archivo. El SAA en TXT no requiere selección de hojas.",
            ),
            (
                "3. Ejecuta CommiFlow",
                "Pulse Ejecutar CommiFlow para iniciar el procesamiento.",
            ),
            (
                "4. Espera a que finalice el proceso",
                "CommiFlow preparará los archivos, procesará las bases VIDA y GMM, el SAA, el Acumulado, VLSP, Catálogo Estatus Pólizas, Catálogos y generará los archivos de salida.",
            ),
            (
                "5. Revisa los resultados",
                "Al finalizar se indicará que el proceso terminó correctamente. Si ocurre un error, se mostrará el mensaje correspondiente.",
            ),
        ]
        for titulo, detalle in pasos:
            ttk.Label(cuerpo, text=titulo, style="CardTitle.TLabel").pack(
                anchor="w", pady=(4, 0)
            )
            ttk.Label(cuerpo, text=detalle, style="Muted.TLabel", wraplength=900).pack(
                anchor="w"
            )

    @staticmethod
    def _marco_hojas(parent, titulo):
        exterior = tk.Frame(parent, bg=COLORES["fondo"])
        exterior.pack(fill="x", padx=28, pady=(0, 4))
        marco = ttk.LabelFrame(exterior, text=titulo, style="Card.TLabelframe")
        marco.pack(fill="x")
        ttk.Label(
            marco,
            text="Seleccione un archivo para ver las hojas disponibles, si aplica.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=6, pady=4)
        return marco

    def _crear_tab_acerca(self):
        contenido = crear_area_desplazable(self.tab_acerca)
        contenedor = tk.Frame(contenido, bg=COLORES["fondo"])
        contenedor.pack(fill="both", expand=True, padx=48, pady=28)

        tarjeta_ext = tk.Frame(contenedor, bg=COLORES["borde"])
        tarjeta_ext.pack(fill="x")
        tarjeta = tk.Frame(tarjeta_ext, bg=COLORES["tarjeta"])
        tarjeta.pack(fill="x", padx=1, pady=1)

        banner = tk.Frame(tarjeta, bg=COLORES["header"])
        banner.pack(fill="x")
        logo = self._cargar_logo(alto=120, destino="_logo_acerca")
        if logo is not None:
            tk.Label(banner, image=logo, bg=COLORES["header"], bd=0).pack(pady=18)

        cuerpo = tk.Frame(tarjeta, bg=COLORES["tarjeta"])
        cuerpo.pack(fill="x", padx=36, pady=(24, 32))

        tk.Label(
            cuerpo,
            text="CommiFlow",
            bg=COLORES["tarjeta"],
            fg=COLORES["texto"],
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            cuerpo,
            text="Automatización y procesamiento de comisiones",
            bg=COLORES["tarjeta"],
            fg=COLORES["muted"],
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(4, 16))
        tk.Label(
            cuerpo,
            text=(
                "CommiFlow es una aplicación desarrollada para MetLife con el "
                "objetivo de automatizar y simplificar el procesamiento de "
                "información relacionada con comisiones, optimizando tareas "
                "operativas y facilitando la generación de resultados."
            ),
            bg=COLORES["tarjeta"],
            fg=COLORES["texto"],
            font=("Segoe UI", 10),
            wraplength=760,
            justify="left",
        ).pack(anchor="w")

        tk.Label(
            cuerpo,
            text="Desarrollado por",
            bg=COLORES["tarjeta"],
            fg=COLORES["muted"],
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w", pady=(22, 2))
        tk.Label(
            cuerpo,
            text="Juan Fernando Leon Medellin",
            bg=COLORES["tarjeta"],
            fg=COLORES["texto"],
            font=("Segoe UI", 11),
        ).pack(anchor="w")

        tk.Label(
            cuerpo,
            text="Propiedad de",
            bg=COLORES["tarjeta"],
            fg=COLORES["muted"],
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w", pady=(16, 2))
        tk.Label(
            cuerpo,
            text="MetLife",
            bg=COLORES["tarjeta"],
            fg=COLORES["texto"],
            font=("Segoe UI", 11),
        ).pack(anchor="w")

        tk.Frame(cuerpo, bg=COLORES["borde"], height=1).pack(fill="x", pady=20)
        tk.Label(
            cuerpo,
            text="Versión: 1.0.0",
            bg=COLORES["tarjeta"],
            fg=COLORES["texto"],
            font=("Segoe UI", 10),
        ).pack(anchor="w")
        tk.Label(
            cuerpo,
            text="© 2026 MetLife. Todos los derechos reservados.",
            bg=COLORES["tarjeta"],
            fg=COLORES["muted"],
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 0))

    def _crear_tab_bases(self):
        ttk.Label(self.tab_bases, text="Bases SAP", style="Section.TLabel").pack(pady=15)
        self.entrada_vida = crear_selector_multiples(
            self.tab_bases,
            "Base VIDA (varios Excel originales o carpeta)",
            lambda: self._seleccionar_archivos("vida"),
            lambda: self._seleccionar_carpeta("vida"),
        )
        self.frame_vida_hojas = ttk.LabelFrame(self.tab_bases, text="Hojas VIDA")
        self.frame_vida_hojas.pack(fill="x", padx=20, pady=10)

        self.entrada_gmm = crear_selector_multiples(
            self.tab_bases,
            "Base GMM (varios Excel originales o carpeta)",
            lambda: self._seleccionar_archivos("gmm"),
            lambda: self._seleccionar_carpeta("gmm"),
        )
        self.frame_gmm_hojas = ttk.LabelFrame(self.tab_bases, text="Hojas GMM")
        self.frame_gmm_hojas.pack(fill="x", padx=20, pady=10)

        ttk.Button(
            self.tab_bases,
            text="Generar Bases SAP",
            command=lambda: self._ejecutar_en_hilo(self._worker_bases),
            width=35,
        ).pack(pady=20)

    def _crear_tab_comisiones(self):
        ttk.Label(
            self.tab_comisiones,
            text="Generación de Comisiones",
            style="Section.TLabel",
        ).pack(pady=15)

        frame_checks = ttk.Frame(self.tab_comisiones)
        frame_checks.pack()
        tk.Checkbutton(
            frame_checks, text="Generar VIDA", variable=self.generar_vida_var
        ).pack(side="left", padx=20)
        tk.Checkbutton(
            frame_checks, text="Generar GMM", variable=self.generar_gmm_var
        ).pack(side="left", padx=20)

        self.entrada_saa = crear_selector_multiples(
            self.tab_comisiones,
            "Archivo SAA (un TXT, o varios Excel)",
            lambda: self._seleccionar_archivos("saa"),
            lambda: self._seleccionar_carpeta("saa"),
        )
        self.frame_saa_hojas = ttk.LabelFrame(self.tab_comisiones, text="Hojas SAA")
        self.frame_saa_hojas.pack(fill="x", padx=20, pady=10)

        self.entrada_manuales = crear_selector_multiples(
            self.tab_comisiones,
            "Acumulado Comisiones (varios Excel originales o carpeta)",
            lambda: self._seleccionar_archivos("manuales"),
            lambda: self._seleccionar_carpeta("manuales"),
        )
        self.frame_manuales_hojas = ttk.LabelFrame(
            self.tab_comisiones, text="Hojas Acumulado Comisiones"
        )
        self.frame_manuales_hojas.pack(fill="x", padx=20, pady=10)

        ttk.Button(
            self.tab_comisiones,
            text="Generar Comisiones",
            command=lambda: self._ejecutar_en_hilo(self._worker_comisiones),
            width=35,
        ).pack(pady=20)

    def _crear_tab_reporte(self):
        ttk.Label(self.tab_reporte, text="Reporte Final", style="Section.TLabel").pack(
            pady=15
        )
        self.entrada_vlsp = crear_selector_archivo(
            self.tab_reporte, "Archivo VLSP", lambda: self._seleccionar_archivo("vlsp")
        )
        self.frame_vlsp_hojas = ttk.LabelFrame(self.tab_reporte, text="Hojas VLSP")
        self.frame_vlsp_hojas.pack(fill="x", padx=20, pady=10)

        self.entrada_tipo = crear_selector_archivo(
            self.tab_reporte,
            "Catálogo Estatus Pólizas",
            lambda: self._seleccionar_archivo("tipo"),
        )
        self.frame_tipo_hojas = ttk.LabelFrame(
            self.tab_reporte, text="Hojas Catálogo Estatus Pólizas"
        )
        self.frame_tipo_hojas.pack(fill="x", padx=20, pady=10)

        self.entrada_catalogos = crear_selector_archivo(
            self.tab_reporte,
            "Archivo unico de Catalogos",
            lambda: self._seleccionar_archivo("catalogos"),
        )
        self.frame_catalogos_hojas = ttk.LabelFrame(
            self.tab_reporte, text="Hojas Catálogos"
        )
        self.frame_catalogos_hojas.pack(fill="x", padx=20, pady=10)

        ttk.Button(
            self.tab_reporte,
            text="Generar Reporte Final",
            command=lambda: self._ejecutar_en_hilo(self._worker_reporte),
            width=35,
        ).pack(pady=20)

    def _crear_footer(self):
        footer = tk.Frame(self.root, bg=COLORES["tarjeta"])
        footer.pack(side="bottom", fill="x")
        tk.Frame(footer, bg=COLORES["borde"], height=1).pack(fill="x")

        interior = tk.Frame(footer, bg=COLORES["tarjeta"])
        interior.pack(fill="x", padx=24, pady=12)

        fila_estado = tk.Frame(interior, bg=COLORES["tarjeta"])
        fila_estado.pack(fill="x")
        self.status_label = tk.Label(
            fila_estado,
            text="Listo para comenzar",
            bg=COLORES["tarjeta"],
            fg=COLORES["texto"],
            font=("Segoe UI Semibold", 10),
        )
        self.status_label.pack(side="left")
        self.percent_label = tk.Label(
            fila_estado,
            text="0 %",
            bg=COLORES["tarjeta"],
            fg=COLORES["primario"],
            font=("Segoe UI Semibold", 12),
        )
        self.percent_label.pack(side="right")

        fila_barra = tk.Frame(interior, bg=COLORES["tarjeta"])
        fila_barra.pack(fill="x", pady=(8, 0))
        ttk.Label(fila_barra, text="0%", style="Muted.TLabel").pack(side="left")
        self.progress_bar = ttk.Progressbar(
            fila_barra,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            style="Commi.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Label(fila_barra, text="100%", style="Muted.TLabel").pack(side="right")

    def _seleccionar_archivo(self, clave):
        tipos = TIPOS_SAA if clave == "saa" else TIPOS_EXCEL
        archivo = filedialog.askopenfilename(filetypes=tipos)
        if not archivo:
            return
        self._asignar_archivos(clave, [archivo])

    def _seleccionar_archivos(self, clave):
        tipos = TIPOS_SAA if clave == "saa" else TIPOS_EXCEL
        seleccion = filedialog.askopenfilenames(filetypes=tipos)
        if not seleccion:
            return
        self._asignar_archivos(clave, list(seleccion))

    def _seleccionar_carpeta(self, clave):
        carpeta = filedialog.askdirectory()
        if not carpeta:
            return
        archivos = listar_excel_en_carpeta(carpeta)
        if not archivos:
            messagebox.showerror(
                "Sin archivos",
                "La carpeta no contiene Excel .xls o .xlsx.",
            )
            return
        self._asignar_archivos(clave, archivos)

    def _asignar_archivos(self, clave, archivos):
        if clave == "saa":
            txts = [a for a in archivos if a.lower().endswith(".txt")]
            excels = [a for a in archivos if not a.lower().endswith(".txt")]
            if txts and excels:
                messagebox.showerror(
                    "SAA",
                    "No combine TXT con Excel. Use un solo TXT, o uno/varios Excel.",
                )
                return
            if len(txts) > 1:
                messagebox.showerror(
                    "SAA",
                    "SAA en TXT debe ser un solo archivo.",
                )
                return

        self.archivos[clave] = archivos if clave in CLAVES_MULTIPLES else archivos[0]
        texto = self._texto_seleccion(archivos)
        entradas = {
            "vida": [self.entrada_vida, self.entrada_pc_vida],
            "gmm": [self.entrada_gmm, self.entrada_pc_gmm],
            "saa": [self.entrada_saa, self.entrada_pc_saa],
            "manuales": [self.entrada_manuales, self.entrada_pc_manuales],
            "vlsp": [self.entrada_vlsp, self.entrada_pc_vlsp],
            "tipo": [self.entrada_tipo, self.entrada_pc_tipo],
            "catalogos": [self.entrada_catalogos, self.entrada_pc_catalogos],
        }
        for entrada in entradas[clave]:
            reemplazar_texto(entrada, texto)

        if clave == "saa" and es_entrada_txt(archivos):
            self._actualizar_hojas(
                clave, [], mensaje="No aplica selección de hojas para archivo TXT."
            )
            return

        excels = [a for a in archivos if not a.lower().endswith(".txt")]
        if len(excels) > 1 and clave in ("vida", "gmm", "manuales"):
            self._actualizar_hojas(
                clave,
                [],
                mensaje=(
                    "Varios archivos: se consolidan todas las hojas de cada Excel, "
                    "como el conversor original."
                ),
            )
            return

        if len(excels) == 1:
            hojas = obtener_hojas(excels[0])
        else:
            reemplazar_texto(
                entradas[clave][1],
                f"Leyendo hojas de {len(excels)} archivos...",
            )
            self.root.update_idletasks()
            hojas = obtener_hojas_union(excels)
            for entrada in entradas[clave]:
                reemplazar_texto(entrada, texto)
        self._actualizar_hojas(clave, hojas)

    @staticmethod
    def _texto_seleccion(archivos):
        if len(archivos) == 1:
            return archivos[0]
        carpeta = Path(archivos[0]).parent
        return f"{len(archivos)} archivos en {carpeta}"

    def _seleccionar_excel(self, clave):
        if clave in CLAVES_MULTIPLES:
            self._seleccionar_archivos(clave)
        else:
            self._seleccionar_archivo(clave)

    def _actualizar_hojas(self, clave, hojas, mensaje=None):
        pares = {
            "vida": (self.frame_vida_hojas, self.frame_pc_vida_hojas, "checks_vida", "checks_pc_vida"),
            "gmm": (self.frame_gmm_hojas, self.frame_pc_gmm_hojas, "checks_gmm", "checks_pc_gmm"),
            "saa": (self.frame_saa_hojas, self.frame_pc_saa_hojas, "checks_saa", "checks_pc_saa"),
            "manuales": (
                self.frame_manuales_hojas,
                self.frame_pc_manuales_hojas,
                "checks_manuales",
                "checks_pc_manuales",
            ),
            "vlsp": (self.frame_vlsp_hojas, self.frame_pc_vlsp_hojas, "checks_vlsp", "checks_pc_vlsp"),
            "tipo": (self.frame_tipo_hojas, self.frame_pc_tipo_hojas, "checks_tipo", "checks_pc_tipo"),
            "catalogos": (
                self.frame_catalogos_hojas,
                self.frame_pc_catalogos_hojas,
                "checks_catalogos",
                "checks_pc_catalogos",
            ),
        }

        frame, frame_pc, attr, attr_pc = pares[clave]

        if not hojas:
            texto = mensaje or "No aplica selección de hojas para archivo TXT."
            self._mensaje_hojas(frame, texto)
            self._mensaje_hojas(frame_pc, texto)
            setattr(self, attr, [])
            setattr(self, attr_pc, [])
            return

        setattr(self, attr, llenar_checks(frame, hojas))
        setattr(self, attr_pc, llenar_checks(frame_pc, hojas))

    @staticmethod
    def _mensaje_hojas(frame, texto):
        for widget in frame.winfo_children():
            widget.destroy()
        ttk.Label(frame, text=texto, style="Muted.TLabel").pack(
            anchor="w", padx=5, pady=5
        )

    @staticmethod
    def _llenar_combos(combos, valores):
        for combo in combos:
            combo["values"] = valores
            if valores:
                combo.set(valores[0])

    def _hojas_seleccionadas(self, clave, proceso_completo=False):
        mapa = {
            "vida": (self.checks_pc_vida, self.checks_vida),
            "gmm": (self.checks_pc_gmm, self.checks_gmm),
            "saa": (self.checks_pc_saa, self.checks_saa),
            "manuales": (self.checks_pc_manuales, self.checks_manuales),
            "vlsp": (self.checks_pc_vlsp, self.checks_vlsp),
            "tipo": (self.checks_pc_tipo, self.checks_tipo),
            "catalogos": (self.checks_pc_catalogos, self.checks_catalogos),
        }
        if clave not in mapa:
            raise ValueError(f"No existen hojas seleccionables para: {clave}")

        checks_pc, checks = mapa[clave]
        checks_usar = checks_pc if proceso_completo else checks

        hojas_seleccionadas = [
            hoja
            for hoja, variable in checks_usar
            if variable.get()
        ]

        if not hojas_seleccionadas:
            archivos = self.archivos.get(clave)
            varios = isinstance(archivos, (list, tuple)) and len(archivos) > 1
            if varios and clave in ("vida", "gmm", "manuales"):
                return []
            nombres = {
                "tipo": "Catálogo Estatus Pólizas",
                "catalogos": "Catálogos",
                "manuales": "Acumulado de Comisiones",
            }
            nombre = nombres.get(clave, clave.upper())
            raise ValueError(
                f"Debe seleccionar al menos una hoja de {nombre}."
            )

        return hojas_seleccionadas

    def actualizar_estado(self, texto, progreso, tipo="info"):
        self.root.after(
            0, lambda t=texto, p=progreso, k=tipo: self._aplicar_estado(t, p, k)
        )

    def _aplicar_estado(self, texto, progreso, tipo="info"):
        if progreso is not None:
            try:
                progreso = int(progreso)
            except (TypeError, ValueError):
                progreso = self.progress_var.get()
            progreso = max(0, min(100, progreso))
            if self._progreso_global:
                progreso = max(self._progreso_minimo, progreso)
                self._progreso_minimo = progreso
            self.progress_var.set(progreso)
            self.percent_label.config(text=f"{progreso} %")

        if texto:
            colores = {
                "ok": COLORES["exito"],
                "error": COLORES["error"],
                "info": COLORES["texto"],
            }
            self.status_label.config(text=texto, fg=colores.get(tipo, COLORES["texto"]))

    def _reportar_etapa(self, inicio, fin):
        def _callback(texto, local=0):
            try:
                local = int(local or 0)
            except (TypeError, ValueError):
                local = 0
            local = max(0, min(100, local))
            valor = int(round(inicio + (local / 100.0) * (fin - inicio)))
            self.actualizar_estado(texto, valor)

        return _callback

    def _ejecutar_en_hilo(self, funcion):
        threading.Thread(target=funcion, daemon=True).start()

    def _set_ejecutando(self, ejecutando):
        if ejecutando:
            self.boton_proceso_completo.config(
                state="disabled",
                bg="#9BB8C7",
                cursor="arrow",
            )
        else:
            self.boton_proceso_completo.config(
                state="normal",
                bg=COLORES["primario"],
                cursor="hand2",
            )

    def _manejar_error(self, error):
        import traceback

        traceback.print_exc()
        mensaje = str(error)
        self.root.after(
            0,
            lambda: self._aplicar_estado(f"Error: {mensaje}", None, "error"),
        )
        self.root.after(0, lambda: messagebox.showerror("Error", mensaje))

    def _worker_bases(
        self,
        mostrar_mensaje=True,
        proceso_completo=False,
        callback=None,
    ):
        try:
            generar_bases_sap(
                self.archivos["vida"],
                self.archivos["gmm"],
                self._hojas_seleccionadas(
                    "vida",
                    proceso_completo=proceso_completo
                ),
                self._hojas_seleccionadas(
                    "gmm",
                    proceso_completo=proceso_completo
                ),
                callback or self.actualizar_estado,
            )

            if mostrar_mensaje:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Proceso terminado",
                        "Bases generadas correctamente."
                    )
                )

        except Exception as error:
            if mostrar_mensaje:
                self._manejar_error(error)
            else:
                raise

    def _worker_comisiones(
        self, mostrar_mensaje=True, proceso_completo=False, callback=None
    ):
        try:
            hojas_saa = None
            if self.archivos["saa"] and not es_entrada_txt(self.archivos["saa"]):
                hojas_saa = self._hojas_seleccionadas(
                    "saa", proceso_completo=proceso_completo
                )

            generar_comisiones(
                self.archivos["saa"],
                self.archivos["manuales"],
                self.generar_vida_var.get(),
                self.generar_gmm_var.get(),
                callback or self.actualizar_estado,
                hojas_saa=hojas_saa,
                hojas_manuales=self._hojas_seleccionadas(
                    "manuales", proceso_completo=proceso_completo
                ),
            )
            if mostrar_mensaje:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Proceso terminado", "Comisiones generadas correctamente."
                    ),
                )
        except Exception as error:
            if mostrar_mensaje:
                self._manejar_error(error)
            else:
                raise

    def _worker_reporte(
        self, mostrar_mensaje=True, proceso_completo=False, callback=None
    ):
        try:
            generar_reporte_final(
                self.archivos["vlsp"],
                self._hojas_seleccionadas("vlsp", proceso_completo=proceso_completo),
                self.archivos["tipo"],
                self._hojas_seleccionadas("tipo", proceso_completo=proceso_completo),
                self.archivos["catalogos"],
                callback or self.actualizar_estado,
                hojas_catalogos=self._hojas_seleccionadas(
                    "catalogos", proceso_completo=proceso_completo
                ),
            )
            if mostrar_mensaje:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Proceso terminado", "Reporte generado correctamente."
                    ),
                )
        except Exception as error:
            if mostrar_mensaje:
                self._manejar_error(error)
            else:
                raise

    def ejecutar_proceso_completo(self):
        self._ejecutar_en_hilo(self._worker_proceso_completo)

    def _worker_proceso_completo(self):
        self.root.after(0, lambda: self._set_ejecutando(True))
        self._progreso_global = True
        self._progreso_minimo = 0
        self.actualizar_estado("Iniciando CommiFlow...", 0)
        try:
            self._worker_bases(
                mostrar_mensaje=False,
                proceso_completo=True,
                callback=self._reportar_etapa(2, 45),
            )
            self._worker_comisiones(
                mostrar_mensaje=False,
                proceso_completo=True,
                callback=self._reportar_etapa(45, 76),
            )
            self._worker_reporte(
                mostrar_mensaje=False,
                proceso_completo=True,
                callback=self._reportar_etapa(76, 98),
            )
            self.actualizar_estado(
                "Proceso completado correctamente", 100, "ok"
            )
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "CommiFlow",
                    "Proceso completado correctamente.",
                ),
            )
        except Exception as error:
            self._manejar_error(error)
        finally:
            self._progreso_global = False
            self.root.after(0, lambda: self._set_ejecutando(False))


def iniciar_app():
    root = tk.Tk()
    VentanaPrincipal(root)
    root.mainloop()


if __name__ == "__main__":
    iniciar_app()
