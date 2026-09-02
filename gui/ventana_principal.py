import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from gui.componentes import (
    crear_area_desplazable,
    crear_selector_archivo,
    llenar_checks,
    reemplazar_texto,
)
from procesos.bases_sap import generar_bases_sap
from procesos.comisiones import generar_comisiones
from procesos.reporte_final import generar_reporte_final
from servicios.excel import obtener_hojas


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

        self.checks_pc_vida = []
        self.checks_pc_gmm = []

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
        self.notebook.add(self.tab_bases, text="Bases SAP")
        self.notebook.add(self.tab_comisiones, text="Comisiones")
        self.notebook.add(self.tab_reporte, text="Reporte Final")

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
            text="Ejecuta Bases SAP, Comisiones y Reporte Final en una sola acción.",
        ).pack(pady=(0, 15))

        self.entrada_pc_vida = crear_selector_archivo(
            contenido, "Base VIDA", lambda: self._seleccionar_excel("vida")
        )
        self.frame_pc_vida_hojas = ttk.LabelFrame(contenido, text="Hojas VIDA")
        self.frame_pc_vida_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_gmm = crear_selector_archivo(
            contenido, "Base GMM", lambda: self._seleccionar_excel("gmm")
        )
        self.frame_pc_gmm_hojas = ttk.LabelFrame(contenido, text="Hojas GMM")
        self.frame_pc_gmm_hojas.pack(fill="x", padx=20, pady=5)

        self.entrada_pc_saa = crear_selector_archivo(
            contenido, "Archivo SAA", lambda: self._seleccionar_excel("saa")
        )
        self.entrada_pc_manuales = crear_selector_archivo(
            contenido, "Acumulado Comisiones", lambda: self._seleccionar_excel("manuales")
        )
        self.entrada_pc_vlsp = crear_selector_archivo(
            contenido, "Archivo VLSP", lambda: self._seleccionar_excel("vlsp")
        )
        ttk.Label(contenido, text="Hoja VLSP").pack(pady=(10, 5))
        self.combo_vlsp_pc = ttk.Combobox(contenido, width=60, state="readonly")
        self.combo_vlsp_pc.pack()

        self.entrada_pc_tipo = crear_selector_archivo(
            contenido, "Archivo Tipo", lambda: self._seleccionar_excel("tipo")
        )
        ttk.Label(contenido, text="Hoja Catálogo Tipo").pack(pady=(10, 5))
        self.combo_tipo_pc = ttk.Combobox(contenido, width=60, state="readonly")
        self.combo_tipo_pc.pack()

        self.entrada_pc_catalogos = crear_selector_archivo(
            contenido, "Archivo unico de Catalogos",
            lambda: self._seleccionar_excel("catalogos"),
        )

        self.boton_proceso_completo = ttk.Button(
            contenido,
            text="Ejecutar Proceso Completo",
            command=self.ejecutar_proceso_completo,
            width=40,
        )
        self.boton_proceso_completo.pack(pady=30)

    def _crear_tab_bases(self):
        ttk.Label(self.tab_bases, text="Bases SAP", style="Header.TLabel").pack(pady=15)
        self.entrada_vida = crear_selector_archivo(
            self.tab_bases, "Base VIDA", lambda: self._seleccionar_excel("vida")
        )
        self.frame_vida_hojas = ttk.LabelFrame(self.tab_bases, text="Hojas VIDA")
        self.frame_vida_hojas.pack(fill="x", padx=20, pady=10)

        self.entrada_gmm = crear_selector_archivo(
            self.tab_bases, "Base GMM", lambda: self._seleccionar_excel("gmm")
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

        self.entrada_saa = crear_selector_archivo(
            self.tab_comisiones, "Archivo SAA", lambda: self._seleccionar_excel("saa")
        )
        self.entrada_manuales = crear_selector_archivo(
            self.tab_comisiones,
            "Acumulado Comisiones",
            lambda: self._seleccionar_excel("manuales"),
        )
        ttk.Button(
            self.tab_comisiones,
            text="Generar Comisiones",
            command=lambda: self._ejecutar_en_hilo(self._worker_comisiones),
            width=35,
        ).pack(pady=20)

    def _crear_tab_reporte(self):
        ttk.Label(self.tab_reporte, text="Reporte Final", style="Header.TLabel").pack(pady=15)
        self.entrada_vlsp = crear_selector_archivo(
            self.tab_reporte, "Archivo VLSP", lambda: self._seleccionar_excel("vlsp")
        )
        ttk.Label(self.tab_reporte, text="Hoja VLSP").pack(pady=(10, 5))
        self.combo_vlsp = ttk.Combobox(self.tab_reporte, width=60, state="readonly")
        self.combo_vlsp.pack()

        self.entrada_tipo = crear_selector_archivo(
            self.tab_reporte, "Archivo Tipo", lambda: self._seleccionar_excel("tipo")
        )
        ttk.Label(self.tab_reporte, text="Hoja Catálogo Tipo").pack(pady=(10, 5))
        self.combo_tipo = ttk.Combobox(self.tab_reporte, width=60, state="readonly")
        self.combo_tipo.pack()

        self.entrada_catalogos = crear_selector_archivo(
            self.tab_reporte, "Archivo unico de Catalogos",
            lambda: self._seleccionar_excel("catalogos"),
        )

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

    def _seleccionar_excel(self, clave):
        archivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not archivo:
            return

        self.archivos[clave] = archivo
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
            reemplazar_texto(entrada, archivo)

        if clave in ("vida", "gmm", "vlsp", "tipo"):
            hojas = obtener_hojas(archivo)
            self._actualizar_hojas(clave, hojas)

    def _actualizar_hojas(self, clave, hojas):
        if clave == "vida":
            self.checks_vida = llenar_checks(
                self.frame_vida_hojas,
                hojas
            )

            self.checks_pc_vida = llenar_checks(
                self.frame_pc_vida_hojas,
                hojas
            )

        elif clave == "gmm":
            self.checks_gmm = llenar_checks(
                self.frame_gmm_hojas,
                hojas
            )

            self.checks_pc_gmm = llenar_checks(
                self.frame_pc_gmm_hojas,
                hojas
            )

        elif clave == "vlsp":
            self._llenar_combos(
                [
                    self.combo_vlsp,
                    self.combo_vlsp_pc
                ],
                hojas
            )

        elif clave == "tipo":
            self._llenar_combos(
                [
                    self.combo_tipo,
                    self.combo_tipo_pc
                ],
                hojas
            )

    @staticmethod
    def _llenar_combos(combos, valores):
        for combo in combos:
            combo["values"] = valores
            if valores:
                combo.set(valores[0])

    def _hojas_seleccionadas(self, clave, proceso_completo=False):
        if clave == "vida":
            checks = (
                self.checks_pc_vida
                if proceso_completo
                else self.checks_vida
            )

        elif clave == "gmm":
            checks = (
                self.checks_pc_gmm
                if proceso_completo
                else self.checks_gmm
            )

        else:
            raise ValueError(
                f"No existen hojas seleccionables para: {clave}"
            )

        hojas_seleccionadas = [
            hoja
            for hoja, variable in checks
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

    def _worker_comisiones(self, mostrar_mensaje=True):
        try:
            generar_comisiones(
                self.archivos["saa"],
                self.archivos["manuales"],
                self.generar_vida_var.get(),
                self.generar_gmm_var.get(),
                self.actualizar_estado,
            )
            if mostrar_mensaje:
                self.root.after(0, lambda: messagebox.showinfo("Proceso terminado", "Comisiones generadas correctamente."))
        except Exception as error:
            if mostrar_mensaje:
                self._manejar_error(error)
            else:
                raise

    def _worker_reporte(self, mostrar_mensaje=True):
        try:
            generar_reporte_final(
                self.archivos["vlsp"],
                self.combo_vlsp_pc.get() or self.combo_vlsp.get(),
                self.archivos["tipo"],
                self.combo_tipo_pc.get() or self.combo_tipo.get(),
                self.archivos["catalogos"],
                self.actualizar_estado,
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
            self._worker_comisiones(mostrar_mensaje=False)
            self._worker_reporte(mostrar_mensaje=False)
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
