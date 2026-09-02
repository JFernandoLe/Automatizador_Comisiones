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
from procesos.bases_sap import generar_bases_sap
from procesos.comisiones import generar_comisiones
from procesos.reporte_final import generar_reporte_final
from servicios.excel import listar_excel_en_carpeta, obtener_hojas, obtener_hojas_union

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

        self.generar_vida_var = tk.BooleanVar(value=True)
        self.generar_gmm_var = tk.BooleanVar(value=True)
        self.progress_var = tk.IntVar(value=0)

        self._configurar_ventana()
        self._crear_interfaz()

    def _configurar_ventana(self):
        self.root.title("Automatizador de Comisiones")
        self.root.geometry("1050x760")
        self.root.minsize(900, 650)
        self.root.configure(bg="#F5F7FA")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=8, font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Horizontal.TProgressbar", thickness=18)

    def _crear_interfaz(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        self.tab_completo = ttk.Frame(self.notebook)
        self.tab_bases = ttk.Frame(self.notebook)
        self.tab_comisiones = ttk.Frame(self.notebook)
        self.tab_reporte = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_completo, text="Proceso Completo")

        self._crear_tab_completo()
        self._crear_tab_bases()
        self._crear_tab_comisiones()
        self._crear_tab_reporte()
        self._crear_footer()

    def _crear_tab_completo(self):
        contenido = crear_area_desplazable(self.tab_completo)
        ttk.Label(contenido, text="Proceso Completo", style="Header.TLabel").pack(pady=(20, 5))
        ttk.Label(
            contenido,
            text=(
                "VIDA, GMM y Acumulado aceptan muchos Excel originales o una carpeta "
                "(como el conversor anterior). SAA acepta un TXT o uno/varios Excel."
            ),
        ).pack(pady=(0, 15))

        self.entrada_pc_vida = crear_selector_multiples(
            contenido,
            "Base VIDA (varios Excel originales o carpeta)",
            lambda: self._seleccionar_archivos("vida"),
            lambda: self._seleccionar_carpeta("vida"),
        )
        self.frame_pc_vida_hojas = ttk.LabelFrame(contenido, text="Hojas VIDA")
        self.frame_pc_vida_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_gmm = crear_selector_multiples(
            contenido,
            "Base GMM (varios Excel originales o carpeta)",
            lambda: self._seleccionar_archivos("gmm"),
            lambda: self._seleccionar_carpeta("gmm"),
        )
        self.frame_pc_gmm_hojas = ttk.LabelFrame(contenido, text="Hojas GMM")
        self.frame_pc_gmm_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_saa = crear_selector_multiples(
            contenido,
            "Archivo SAA (un TXT, o varios Excel)",
            lambda: self._seleccionar_archivos("saa"),
            lambda: self._seleccionar_carpeta("saa"),
        )
        self.frame_pc_saa_hojas = ttk.LabelFrame(contenido, text="Hojas SAA")
        self.frame_pc_saa_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_manuales = crear_selector_multiples(
            contenido,
            "Acumulado Comisiones (varios Excel originales o carpeta)",
            lambda: self._seleccionar_archivos("manuales"),
            lambda: self._seleccionar_carpeta("manuales"),
        )
        self.frame_pc_manuales_hojas = ttk.LabelFrame(
            contenido, text="Hojas Acumulado Comisiones"
        )
        self.frame_pc_manuales_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_vlsp = crear_selector_archivo(
            contenido, "Archivo VLSP", lambda: self._seleccionar_archivo("vlsp")
        )
        self.frame_pc_vlsp_hojas = ttk.LabelFrame(contenido, text="Hojas VLSP")
        self.frame_pc_vlsp_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_tipo = crear_selector_archivo(
            contenido, "Archivo Tipo", lambda: self._seleccionar_archivo("tipo")
        )
        self.frame_pc_tipo_hojas = ttk.LabelFrame(contenido, text="Hojas Catálogo Tipo")
        self.frame_pc_tipo_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_catalogos = crear_selector_archivo(
            contenido,
            "Archivo unico de Catalogos",
            lambda: self._seleccionar_archivo("catalogos"),
        )
        self.frame_pc_catalogos_hojas = ttk.LabelFrame(
            contenido, text="Hojas Catálogos"
        )
        self.frame_pc_catalogos_hojas.pack(fill="x", padx=20, pady=5)

        self.boton_proceso_completo = ttk.Button(
            contenido,
            text="Ejecutar Proceso Completo",
            command=self.ejecutar_proceso_completo,
            width=40,
        )
        self.boton_proceso_completo.pack(pady=30)

    def _crear_tab_bases(self):
        ttk.Label(self.tab_bases, text="Bases SAP", style="Header.TLabel").pack(pady=15)
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
            style="Header.TLabel",
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
        ttk.Label(self.tab_reporte, text="Reporte Final", style="Header.TLabel").pack(pady=15)
        self.entrada_vlsp = crear_selector_archivo(
            self.tab_reporte, "Archivo VLSP", lambda: self._seleccionar_archivo("vlsp")
        )
        self.frame_vlsp_hojas = ttk.LabelFrame(self.tab_reporte, text="Hojas VLSP")
        self.frame_vlsp_hojas.pack(fill="x", padx=20, pady=10)

        self.entrada_tipo = crear_selector_archivo(
            self.tab_reporte, "Archivo Tipo", lambda: self._seleccionar_archivo("tipo")
        )
        self.frame_tipo_hojas = ttk.LabelFrame(
            self.tab_reporte, text="Hojas Catálogo Tipo"
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
        footer = ttk.Frame(self.root)
        footer.pack(fill="x", side="bottom", padx=10, pady=10)
        ttk.Progressbar(
            footer,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        ).pack(fill="x", pady=(0, 5))
        self.status_label = ttk.Label(footer, text="Listo", font=("Segoe UI", 10, "bold"))
        self.status_label.pack(anchor="w")

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
            self._actualizar_hojas(clave, [])
            return

        excels = [a for a in archivos if not a.lower().endswith(".txt")]
        if len(excels) == 1:
            hojas = obtener_hojas(excels[0])
        else:
            reemplazar_texto(entradas[clave][1], f"Leyendo hojas de {len(excels)} archivos...")
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

    def _actualizar_hojas(self, clave, hojas):
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
            mensaje = "No aplica selección de hojas para archivo TXT."
            self._mensaje_hojas(frame, mensaje)
            self._mensaje_hojas(frame_pc, mensaje)
            setattr(self, attr, [])
            setattr(self, attr_pc, [])
            return

        setattr(self, attr, llenar_checks(frame, hojas))
        setattr(self, attr_pc, llenar_checks(frame_pc, hojas))

    @staticmethod
    def _mensaje_hojas(frame, texto):
        for widget in frame.winfo_children():
            widget.destroy()
        ttk.Label(frame, text=texto).pack(anchor="w", padx=5, pady=5)

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
            raise ValueError(
                f"Debe seleccionar al menos una hoja de {clave.upper()}."
            )

        return hojas_seleccionadas

    def actualizar_estado(self, texto, progreso):
        self.root.after(0, lambda: self._aplicar_estado(texto, progreso))

    def _aplicar_estado(self, texto, progreso):
        self.status_label.config(text=f"Estado: {texto}")
        self.progress_var.set(progreso)

    def _ejecutar_en_hilo(self, funcion):
        threading.Thread(target=funcion, daemon=True).start()

    def _manejar_error(self, error):
        import traceback
        traceback.print_exc()
        self.root.after(0, lambda: messagebox.showerror("Error", str(error)))

    def _worker_bases(
        self,
        mostrar_mensaje=True,
        proceso_completo=False
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
                self.actualizar_estado,
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

    def _worker_comisiones(self, mostrar_mensaje=True, proceso_completo=False):
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
                self.actualizar_estado,
                hojas_saa=hojas_saa,
                hojas_manuales=self._hojas_seleccionadas(
                    "manuales", proceso_completo=proceso_completo
                ),
            )
            if mostrar_mensaje:
                self.root.after(0, lambda: messagebox.showinfo("Proceso terminado", "Comisiones generadas correctamente."))
        except Exception as error:
            if mostrar_mensaje:
                self._manejar_error(error)
            else:
                raise

    def _worker_reporte(self, mostrar_mensaje=True, proceso_completo=False):
        try:
            generar_reporte_final(
                self.archivos["vlsp"],
                self._hojas_seleccionadas("vlsp", proceso_completo=proceso_completo),
                self.archivos["tipo"],
                self._hojas_seleccionadas("tipo", proceso_completo=proceso_completo),
                self.archivos["catalogos"],
                self.actualizar_estado,
                hojas_catalogos=self._hojas_seleccionadas(
                    "catalogos", proceso_completo=proceso_completo
                ),
            )
            if mostrar_mensaje:
                self.root.after(0, lambda: messagebox.showinfo("Proceso terminado", "Reporte generado correctamente."))
        except Exception as error:
            if mostrar_mensaje:
                self._manejar_error(error)
            else:
                raise
            
    def ejecutar_proceso_completo(self):
        self._ejecutar_en_hilo(self._worker_proceso_completo)

    def _worker_proceso_completo(self):
        self.root.after(0, lambda: self.boton_proceso_completo.config(state="disabled"))
        try:
            self._worker_bases(
                mostrar_mensaje=False,
                proceso_completo=True
            )
            self._worker_comisiones(
                mostrar_mensaje=False,
                proceso_completo=True,
            )
            self._worker_reporte(
                mostrar_mensaje=False,
                proceso_completo=True,
            )
            self.root.after(0, lambda: messagebox.showinfo("Proceso terminado", "Proceso completo ejecutado correctamente."))
        except Exception as error:
            self._manejar_error(error)
        finally:
            self.root.after(0, lambda: self.boton_proceso_completo.config(state="normal"))


def iniciar_app():
    root = tk.Tk()
    VentanaPrincipal(root)
    root.mainloop()


if __name__ == "__main__":
    iniciar_app()
