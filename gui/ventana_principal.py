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
from fase2.bases_pp import (
    ORDEN_TIPOS_PP,
    TITULOS_PP,
    agrupar_excel_pp,
    extraer_mes_anio,
    generar_primas,
    listar_excel_primas,
)
from fase2.procesador import generar_comision_fase2
from fase2.transformaciones import (
    CONCEPTOS_DETALLE,
    FIGURAS_DETALLE,
    RAMOS2_DETALLE,
)
from procesos.bases_sap import generar_bases_sap
from procesos.comisiones import generar_comisiones
from procesos.reporte_final import generar_reporte_final
from servicios.excel import listar_excel_en_carpeta, obtener_hojas, obtener_hojas_union
from servicios.recursos import ruta_icono, ruta_logo

TIPOS_EXCEL = [("Excel", "*.xlsx *.xls")]
TIPOS_MANUALES = [
    ("Excel", "*.xlsx *.xls *.xlsb"),
    ("Excel binario", "*.xlsb"),
]
TIPOS_SAA = [
    ("SAA", "*.txt *.xlsx *.xls"),
    ("Texto", "*.txt"),
    ("Excel", "*.xlsx *.xls"),
]
TIPOS_REPORTE = [
    ("Reportes", "*.xlsx *.xls *.parquet"),
    ("Excel", "*.xlsx *.xls"),
    ("Parquet", "*.parquet"),
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
        self.archivos_fase2 = {
            "vida": None,
            "gmm": None,
            "dist": None,
            "catalogos": None,
            "bonos": None,
            "clasif": None,
        }
        self.checks_f2_vida = []
        self.checks_f2_gmm = []
        self.checks_f2_dist = []
        self.checks_f2_catalogos = []
        self.checks_f2_bonos = []
        self.checks_f2_clasif = []
        self.checks_f2_figuras = []
        self.checks_f2_ramos = []
        self.checks_f2_conceptos = []
        self.carpetas_pp = {"actual": None, "anterior": None}
        self.entradas_pp = {}
        self.pp_tipos = {"actual": {}, "anterior": {}}
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
        self._icono_ventana = None

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
        self._aplicar_icono()

    def _aplicar_icono(self):
        ruta = ruta_icono()
        if ruta is None:
            return
        ruta_abs = str(ruta.resolve())
        try:
            self.root.iconbitmap(default=ruta_abs)
            self.root.iconbitmap(ruta_abs)
        except Exception:
            pass
        try:
            from PIL import Image, ImageTk

            imagen = Image.open(ruta)
            self._icono_ventana = ImageTk.PhotoImage(imagen)
            self.root.iconphoto(True, self._icono_ventana)
        except Exception:
            pass

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
        self.tab_fase2 = ttk.Frame(self.notebook, style="Fondo.TFrame")
        self.tab_bases_pp = ttk.Frame(self.notebook, style="Fondo.TFrame")
        self.tab_acerca = ttk.Frame(self.notebook, style="Fondo.TFrame")
        self.notebook.add(self.tab_completo, text="Proceso")
        self.notebook.add(self.tab_fase2, text="Fase 2")
        self.notebook.add(self.tab_bases_pp, text="Bases PP")
        self.notebook.add(self.tab_acerca, text="Acerca de")

        self.tab_bases = ttk.Frame(self.root)
        self.tab_comisiones = ttk.Frame(self.root)
        self.tab_reporte = ttk.Frame(self.root)

        self._crear_tab_completo()
        self._crear_tab_fase2()
        self._crear_tab_bases_pp()
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
            ayuda="Uno o varios Excel (.xlsx, .xls o .xlsb), o una carpeta.",
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

    def _crear_checks_combinaciones(self, parent):
        exterior = tk.Frame(parent, bg=COLORES["fondo"])
        exterior.pack(fill="x", padx=28, pady=(0, 8))
        grupos = (
            ("Figura", FIGURAS_DETALLE, "checks_f2_figuras", {"PROMOTOR", "AGENTE"}),
            ("Ramo", RAMOS2_DETALLE, "checks_f2_ramos", set(RAMOS2_DETALLE)),
            ("Ini_Ren", CONCEPTOS_DETALLE, "checks_f2_conceptos", set(CONCEPTOS_DETALLE)),
        )
        for titulo, valores, attr, activos in grupos:
            marco = ttk.LabelFrame(exterior, text=titulo, style="Card.TLabelframe")
            marco.pack(side="left", fill="x", expand=True, padx=(0, 8))
            checks = []
            for valor in valores:
                variable = tk.BooleanVar(value=valor in activos)
                ttk.Checkbutton(marco, text=valor, variable=variable).pack(
                    anchor="w", padx=6, pady=2
                )
                checks.append((valor, variable))
            setattr(self, attr, checks)

    def _combinaciones_fase2(self):
        def seleccion(checks, nombre):
            valores = [valor for valor, variable in checks if variable.get()]
            if not valores:
                raise ValueError(f"Debe seleccionar al menos un valor de {nombre}.")
            return valores

        return {
            "figuras": seleccion(self.checks_f2_figuras, "Figura"),
            "ramos": seleccion(self.checks_f2_ramos, "Ramo"),
            "conceptos": seleccion(self.checks_f2_conceptos, "Ini_Ren"),
        }

    def _crear_tab_fase2(self):
        contenido = crear_area_desplazable(self.tab_fase2)

        exterior = tk.Frame(contenido, bg=COLORES["fondo"])
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
            cuerpo, text="Fase 2 · Comisión", style="CardTitle.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            cuerpo,
            text=(
                "Usa los reportes finales de la Fase 1, Distribución Comercial, "
                "catálogos y Bonos. Puede ejecutarse sin repetir la Fase 1 si "
                "esos reportes ya existen. El resultado se guarda como Comision.xlsx "
                "con la hoja Pagos_de_Bonos."
            ),
            style="Muted.TLabel",
            wraplength=900,
        ).pack(anchor="w", pady=(4, 0))

        ttk.Label(contenido, text="Reportes de la Fase 1", style="Section.TLabel").pack(
            anchor="w", padx=28, pady=(8, 2)
        )

        self.entrada_f2_vida = crear_selector_archivo(
            contenido,
            "Reporte VIDA Final",
            lambda: self._seleccionar_archivo_fase2("vida"),
            ayuda="Excel o parquet generado por la Fase 1.",
        )
        self.frame_f2_vida_hojas = self._marco_hojas(
            contenido, "Hojas de Reporte VIDA Final"
        )

        self.entrada_f2_gmm = crear_selector_archivo(
            contenido,
            "Reporte GMM Final",
            lambda: self._seleccionar_archivo_fase2("gmm"),
            ayuda="Excel o parquet generado por la Fase 1.",
        )
        self.frame_f2_gmm_hojas = self._marco_hojas(
            contenido, "Hojas de Reporte GMM Final"
        )

        ttk.Label(
            contenido, text="Referencias de la Fase 2", style="Section.TLabel"
        ).pack(anchor="w", padx=28, pady=(18, 2))

        self.entrada_f2_dist = crear_selector_archivo(
            contenido,
            "Distribución Comercial",
            lambda: self._seleccionar_archivo_fase2("dist"),
            ayuda="Excel. Se busca Promotor en la columna A y se toma la columna B.",
        )
        self.frame_f2_dist_hojas = self._marco_hojas(
            contenido, "Hojas de Distribución Comercial"
        )

        self.entrada_f2_catalogos = crear_selector_archivo(
            contenido,
            "Archivo de Catálogos",
            lambda: self._seleccionar_archivo_fase2("catalogos"),
            ayuda="Debe incluir la hoja PFPM (AGENTE_ORIGINAL y AGENTE_REPORTERIA).",
        )
        self.frame_f2_catalogos_hojas = self._marco_hojas(
            contenido, "Hojas de Catálogos"
        )

        self.entrada_f2_bonos = crear_selector_archivo(
            contenido,
            "Bonos",
            lambda: self._seleccionar_archivo_fase2("bonos"),
            ayuda="Excel. Encabezados en la fila 2; los datos empiezan en la fila 3.",
        )
        self.frame_f2_bonos_hojas = self._marco_hojas(contenido, "Hojas de Bonos")

        self.entrada_f2_clasif = crear_selector_archivo(
            contenido,
            "Catálogo de clasificaciones",
            lambda: self._seleccionar_archivo_fase2("clasif"),
            ayuda="Excel con hoja VIDA y hoja GMM. Encabezados: Ramo, Ini_Ren, FIGURA, CATEGORIA, COMPENSACIÓN, CC, NOMBRE.",
        )
        self.frame_f2_clasif_hojas = self._marco_hojas(
            contenido, "Hojas de Catálogo de clasificaciones"
        )

        ttk.Label(
            contenido,
            text="Tablas de combinaciones",
            style="Section.TLabel",
        ).pack(anchor="w", padx=28, pady=(18, 2))
        ttk.Label(
            contenido,
            text="Marque figura, ramo e Ini_Ren. Se genera una tabla por cada combinación con venta distinta de 0.",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=28, pady=(0, 8))
        self._crear_checks_combinaciones(contenido)

        boton_frame = tk.Frame(contenido, bg=COLORES["fondo"])
        boton_frame.pack(pady=(24, 36))
        self.boton_fase2 = tk.Button(
            boton_frame,
            text="Generar Comisión",
            command=self.ejecutar_fase2,
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
        self.boton_fase2.pack()
        self.boton_fase2.bind(
            "<Enter>",
            lambda _e: self.boton_fase2.config(bg=COLORES["primario_hover"])
            if str(self.boton_fase2["state"]) == "normal"
            else None,
        )
        self.boton_fase2.bind(
            "<Leave>",
            lambda _e: self.boton_fase2.config(bg=COLORES["primario"])
            if str(self.boton_fase2["state"]) == "normal"
            else None,
        )

    def _crear_tab_bases_pp(self):
        contenido = crear_area_desplazable(self.tab_bases_pp)

        exterior = tk.Frame(contenido, bg=COLORES["fondo"])
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
            cuerpo, text="Bases PP · Primas ponderadas / pagadas", style="CardTitle.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            cuerpo,
            text=(
                "Seleccione las carpetas Bases de Primas del mes del año actual y "
                "del año anterior. Solo se leen Excel del primer nivel. Confirme "
                "los 7 archivos de cada año y la hoja de datos. El resultado se "
                "guarda como Primas.xlsx, hoja Bases PP. Etapa independiente para pruebas."
            ),
            style="Muted.TLabel",
            wraplength=900,
        ).pack(anchor="w", pady=(4, 0))

        self.entradas_pp["actual"] = self._crear_seccion_pp(
            contenido,
            "actual",
            "Año actual",
            "Carpeta Bases de Primas del mes (año actual)",
        )
        self.entradas_pp["anterior"] = self._crear_seccion_pp(
            contenido,
            "anterior",
            "Año anterior",
            "Carpeta Bases de Primas del mes (año anterior)",
        )

        boton_frame = tk.Frame(contenido, bg=COLORES["fondo"])
        boton_frame.pack(pady=(24, 36))
        self.boton_bases_pp = tk.Button(
            boton_frame,
            text="Generar Primas",
            command=self.ejecutar_bases_pp,
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
        self.boton_bases_pp.pack()
        self.boton_bases_pp.bind(
            "<Enter>",
            lambda _e: self.boton_bases_pp.config(bg=COLORES["primario_hover"])
            if str(self.boton_bases_pp["state"]) == "normal"
            else None,
        )
        self.boton_bases_pp.bind(
            "<Leave>",
            lambda _e: self.boton_bases_pp.config(bg=COLORES["primario"])
            if str(self.boton_bases_pp["state"]) == "normal"
            else None,
        )

    def _crear_seccion_pp(self, parent, periodo, titulo_seccion, titulo_selector):
        ttk.Label(parent, text=titulo_seccion, style="Section.TLabel").pack(
            anchor="w", padx=28, pady=(18, 2)
        )
        entrada = crear_selector_archivo(
            parent,
            titulo_selector,
            lambda: self._seleccionar_carpeta_pp(periodo),
            ayuda="Se ignoran subcarpetas, PDF y cualquier archivo que no sea Excel.",
        )
        ttk.Label(
            parent,
            text="Marque el archivo de cada tipo y la hoja donde están Promotor, Agente, Prima Pagada y Mes.",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=28, pady=(0, 8))
        for tipo_id in ORDEN_TIPOS_PP:
            self._crear_fila_tipo_pp(parent, periodo, tipo_id)
        return entrada

    def _crear_fila_tipo_pp(self, parent, periodo, tipo_id):
        exterior = tk.Frame(parent, bg=COLORES["fondo"])
        exterior.pack(fill="x", padx=28, pady=(0, 6))
        marco = ttk.LabelFrame(
            exterior, text=TITULOS_PP[tipo_id], style="Card.TLabelframe"
        )
        marco.pack(fill="x")
        activo = tk.BooleanVar(value=False)
        ttk.Checkbutton(marco, text="Usar este archivo", variable=activo).pack(
            anchor="w", padx=6, pady=(4, 0)
        )
        combo = ttk.Combobox(marco, state="readonly", width=88)
        combo.pack(fill="x", padx=6, pady=4)
        combo.bind(
            "<<ComboboxSelected>>",
            lambda _e, p=periodo, t=tipo_id: self._cambiar_archivo_pp(p, t),
        )
        frame_hojas = tk.Frame(marco, bg=COLORES["tarjeta"])
        frame_hojas.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Label(
            frame_hojas,
            text="Seleccione una carpeta para listar Excel y hojas.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=4, pady=4)
        self.pp_tipos[periodo][tipo_id] = {
            "activo": activo,
            "combo": combo,
            "frame_hojas": frame_hojas,
            "checks": [],
            "rutas": {},
        }

    def _seleccionar_carpeta_pp(self, periodo):
        carpeta = filedialog.askdirectory()
        if not carpeta:
            return
        try:
            archivos = listar_excel_primas(carpeta)
        except ValueError as error:
            messagebox.showerror("Bases PP", str(error))
            return
        self.carpetas_pp[periodo] = carpeta
        reemplazar_texto(self.entradas_pp[periodo], carpeta)
        self._poblar_tipos_pp(periodo, archivos)

    def _poblar_tipos_pp(self, periodo, archivos):
        grupos, _sin_tipo = agrupar_excel_pp(archivos)
        nombres = {Path(archivo).name: archivo for archivo in archivos}
        valores = list(nombres.keys())
        for tipo_id in ORDEN_TIPOS_PP:
            widgets = self.pp_tipos[periodo][tipo_id]
            widgets["rutas"] = nombres
            widgets["combo"]["values"] = valores
            coincidencias = grupos.get(tipo_id) or []
            if coincidencias:
                elegido = Path(coincidencias[0]).name
                widgets["combo"].set(elegido)
                widgets["activo"].set(True)
                self._cargar_hojas_pp(periodo, tipo_id)
            else:
                widgets["combo"].set("")
                widgets["activo"].set(False)
                self._mensaje_hojas(
                    widgets["frame_hojas"],
                    "No se detectó este tipo. Elija el Excel si aplica.",
                )
                widgets["checks"] = []

    def _cambiar_archivo_pp(self, periodo, tipo_id):
        widgets = self.pp_tipos[periodo][tipo_id]
        if widgets["combo"].get():
            widgets["activo"].set(True)
            self._cargar_hojas_pp(periodo, tipo_id)

    def _cargar_hojas_pp(self, periodo, tipo_id):
        widgets = self.pp_tipos[periodo][tipo_id]
        nombre = widgets["combo"].get()
        ruta = widgets["rutas"].get(nombre)
        if not ruta:
            self._mensaje_hojas(
                widgets["frame_hojas"], "Seleccione el Excel de este tipo."
            )
            widgets["checks"] = []
            return
        hojas = obtener_hojas(ruta)
        if not hojas:
            self._mensaje_hojas(widgets["frame_hojas"], "El archivo no contiene hojas.")
            widgets["checks"] = []
            return
        widgets["checks"] = llenar_checks(widgets["frame_hojas"], hojas)
        if len(hojas) > 1:
            for _hoja, variable in widgets["checks"]:
                variable.set(False)

    def _entradas_bases_pp(self):
        entradas = []
        faltantes = []
        periodos = (("actual", "Año actual"), ("anterior", "Año anterior"))
        for periodo, etiqueta in periodos:
            if not self.carpetas_pp[periodo]:
                raise ValueError(f"Debe seleccionar la carpeta de {etiqueta}.")
            for tipo_id in ORDEN_TIPOS_PP:
                widgets = self.pp_tipos[periodo][tipo_id]
                titulo = TITULOS_PP[tipo_id]
                nombre = widgets["combo"].get().strip()
                ruta = widgets["rutas"].get(nombre) if nombre else None
                if not widgets["activo"].get() or not ruta:
                    faltantes.append(f"{etiqueta} · {titulo}")
                    continue
                hojas = [hoja for hoja, variable in widgets["checks"] if variable.get()]
                if not hojas:
                    raise ValueError(
                        f"Debe seleccionar la hoja de datos de {etiqueta} · {titulo}."
                    )
                mes, anio = extraer_mes_anio(ruta)
                entradas.append(
                    {
                        "tipo_id": tipo_id,
                        "titulo": titulo,
                        "anio": anio,
                        "mes_hasta": mes,
                        "archivo": ruta,
                        "hojas": hojas,
                    }
                )
        return entradas, faltantes

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
            "Acumulado Comisiones (varios Excel .xlsx, .xls o .xlsb, o carpeta)",
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

    @staticmethod
    def _tipos_dialogo(clave):
        if clave == "saa":
            return TIPOS_SAA
        if clave == "manuales":
            return TIPOS_MANUALES
        return TIPOS_EXCEL

    def _seleccionar_archivo(self, clave):
        archivo = filedialog.askopenfilename(filetypes=self._tipos_dialogo(clave))
        if not archivo:
            return
        self._asignar_archivos(clave, [archivo])

    def _seleccionar_archivos(self, clave):
        seleccion = filedialog.askopenfilenames(
            filetypes=self._tipos_dialogo(clave)
        )
        if not seleccion:
            return
        self._asignar_archivos(clave, list(seleccion))

    def _seleccionar_carpeta(self, clave):
        carpeta = filedialog.askdirectory()
        if not carpeta:
            return
        archivos = listar_excel_en_carpeta(
            carpeta, incluir_xlsb=(clave == "manuales")
        )
        if not archivos:
            mensaje = (
                "La carpeta no contiene Excel .xls, .xlsx o .xlsb."
                if clave == "manuales"
                else "La carpeta no contiene Excel .xls o .xlsx."
            )
            messagebox.showerror("Sin archivos", mensaje)
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
        botones = [self.boton_proceso_completo]
        if getattr(self, "boton_fase2", None) is not None:
            botones.append(self.boton_fase2)
        if getattr(self, "boton_bases_pp", None) is not None:
            botones.append(self.boton_bases_pp)
        if ejecutando:
            for boton in botones:
                boton.config(state="disabled", bg="#9BB8C7", cursor="arrow")
        else:
            for boton in botones:
                boton.config(
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

    def _seleccionar_archivo_fase2(self, clave):
        tipos = TIPOS_REPORTE if clave in ("vida", "gmm") else TIPOS_EXCEL
        archivo = filedialog.askopenfilename(filetypes=tipos)
        if not archivo:
            return
        self.archivos_fase2[clave] = archivo
        entradas = {
            "vida": self.entrada_f2_vida,
            "gmm": self.entrada_f2_gmm,
            "dist": self.entrada_f2_dist,
            "catalogos": self.entrada_f2_catalogos,
            "bonos": self.entrada_f2_bonos,
            "clasif": self.entrada_f2_clasif,
        }
        reemplazar_texto(entradas[clave], archivo)

        frames = {
            "vida": ("frame_f2_vida_hojas", "checks_f2_vida"),
            "gmm": ("frame_f2_gmm_hojas", "checks_f2_gmm"),
            "dist": ("frame_f2_dist_hojas", "checks_f2_dist"),
            "catalogos": ("frame_f2_catalogos_hojas", "checks_f2_catalogos"),
            "bonos": ("frame_f2_bonos_hojas", "checks_f2_bonos"),
            "clasif": ("frame_f2_clasif_hojas", "checks_f2_clasif"),
        }
        frame_attr, checks_attr = frames[clave]
        frame = getattr(self, frame_attr)

        if Path(archivo).suffix.lower() == ".parquet":
            self._mensaje_hojas(frame, "No aplica selección de hojas para archivo parquet.")
            setattr(self, checks_attr, [])
            return

        self._actualizar_hojas_fase2(clave, obtener_hojas(archivo))

    def _actualizar_hojas_fase2(self, clave, hojas):
        pares = {
            "vida": ("frame_f2_vida_hojas", "checks_f2_vida"),
            "gmm": ("frame_f2_gmm_hojas", "checks_f2_gmm"),
            "dist": ("frame_f2_dist_hojas", "checks_f2_dist"),
            "catalogos": ("frame_f2_catalogos_hojas", "checks_f2_catalogos"),
            "bonos": ("frame_f2_bonos_hojas", "checks_f2_bonos"),
            "clasif": ("frame_f2_clasif_hojas", "checks_f2_clasif"),
        }
        frame_attr, checks_attr = pares[clave]
        frame = getattr(self, frame_attr)
        if not hojas:
            self._mensaje_hojas(frame, "El archivo no contiene hojas.")
            setattr(self, checks_attr, [])
            return
        setattr(self, checks_attr, llenar_checks(frame, hojas))

    def _hojas_fase2(self, clave):
        nombres = {
            "vida": ("Reporte VIDA Final", self.checks_f2_vida),
            "gmm": ("Reporte GMM Final", self.checks_f2_gmm),
            "dist": ("Distribución Comercial", self.checks_f2_dist),
            "catalogos": ("Catálogos", self.checks_f2_catalogos),
            "bonos": ("Bonos", self.checks_f2_bonos),
            "clasif": ("Catálogo de clasificaciones", self.checks_f2_clasif),
        }
        nombre, checks = nombres[clave]
        archivo = self.archivos_fase2.get(clave)
        if archivo and Path(archivo).suffix.lower() == ".parquet":
            return None
        hojas = [hoja for hoja, variable in checks if variable.get()]
        if not hojas:
            raise ValueError(f"Debe seleccionar al menos una hoja de {nombre}.")
        return hojas

    def ejecutar_fase2(self):
        self._ejecutar_en_hilo(self._worker_fase2)

    def _worker_fase2(self):
        self.root.after(0, lambda: self._set_ejecutando(True))
        self._progreso_global = True
        self._progreso_minimo = 0
        self.actualizar_estado("Iniciando Fase 2...", 0)
        try:
            generar_comision_fase2(
                self.archivos_fase2["vida"],
                self._hojas_fase2("vida"),
                self.archivos_fase2["gmm"],
                self._hojas_fase2("gmm"),
                self.archivos_fase2["dist"],
                self._hojas_fase2("dist"),
                self.archivos_fase2["catalogos"],
                self._hojas_fase2("catalogos"),
                self.archivos_fase2["bonos"],
                self._hojas_fase2("bonos"),
                self.archivos_fase2["clasif"],
                self._hojas_fase2("clasif"),
                combinaciones=self._combinaciones_fase2(),
                actualizar_estado=self.actualizar_estado,
            )
            self.actualizar_estado("Fase 2 completada correctamente", 100, "ok")
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "CommiFlow",
                    "Fase 2 completada. Se generó Comision.xlsx con la hoja Pagos_de_Bonos.",
                ),
            )
        except Exception as error:
            self._manejar_error(error)
        finally:
            self._progreso_global = False
            self.root.after(0, lambda: self._set_ejecutando(False))

    def ejecutar_bases_pp(self):
        try:
            entradas, faltantes = self._entradas_bases_pp()
        except ValueError as error:
            messagebox.showerror("Bases PP", str(error))
            return
        if faltantes:
            listado = "\n".join(f"• {item}" for item in faltantes)
            continuar = messagebox.askyesno(
                "Archivos faltantes",
                f"Faltan {len(faltantes)} archivo(s):\n\n{listado}\n\n"
                "¿Desea continuar con los archivos disponibles?",
            )
            if not continuar:
                return
        if not entradas:
            messagebox.showerror(
                "Bases PP",
                "No hay archivos seleccionados para generar Primas.xlsx.",
            )
            return
        self._ejecutar_en_hilo(lambda: self._worker_bases_pp(entradas))

    def _worker_bases_pp(self, entradas):
        self.root.after(0, lambda: self._set_ejecutando(True))
        self._progreso_global = True
        self._progreso_minimo = 0
        self.actualizar_estado("Iniciando Bases PP...", 0)
        try:
            generar_primas(entradas, actualizar_estado=self.actualizar_estado)
            self.actualizar_estado("Bases PP completada correctamente", 100, "ok")
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "CommiFlow",
                    "Bases PP completada. Se generó Primas.xlsx con la hoja Bases PP.",
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
