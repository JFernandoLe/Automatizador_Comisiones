from pathlib import Path

ruta = Path('gui/ventana_principal.py')
texto = ruta.read_text(encoding='utf-8')

cambios = [
    ('            "tipo": None,\n', '            "tipo": None,\n            "catalogos": None,\n'),
    ('        self.combo_tipo_pc.pack()\n\n        self.boton_proceso_completo',
     '        self.combo_tipo_pc.pack()\n\n        self.entrada_pc_catalogos = crear_selector_archivo(\n            contenido, "Archivo unico de Catalogos",\n            lambda: self._seleccionar_excel("catalogos"),\n        )\n\n        self.boton_proceso_completo'),
    ('        self.combo_tipo.pack()\n\n        ttk.Button(',
     '        self.combo_tipo.pack()\n\n        self.entrada_catalogos = crear_selector_archivo(\n            self.tab_reporte, "Archivo unico de Catalogos",\n            lambda: self._seleccionar_excel("catalogos"),\n        )\n\n        ttk.Button('),
    ('            "tipo": [self.entrada_tipo, self.entrada_pc_tipo],\n',
     '            "tipo": [self.entrada_tipo, self.entrada_pc_tipo],\n            "catalogos": [self.entrada_catalogos, self.entrada_pc_catalogos],\n'),
    ('                self.combo_tipo_pc.get() or self.combo_tipo.get(),\n                self.actualizar_estado,',
     '                self.combo_tipo_pc.get() or self.combo_tipo.get(),\n                self.archivos["catalogos"],\n                self.actualizar_estado,'),
]

for anterior, nuevo in cambios:
    if anterior not in texto:
        raise RuntimeError('No se encontro un bloque esperado. No se modifico el archivo.')
    texto = texto.replace(anterior, nuevo, 1)

respaldo = ruta.with_name('ventana_principal_antes_catalogos.py')
respaldo.write_text(ruta.read_text(encoding='utf-8'), encoding='utf-8')
ruta.write_text(texto, encoding='utf-8')
print('GUI actualizada. Respaldo:', respaldo)
