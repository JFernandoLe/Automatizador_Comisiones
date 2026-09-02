import json
from pathlib import Path


def cargar_configuracion():
    ruta = Path(__file__).resolve().parent / "configuracion.json"
    with open(ruta, "r", encoding="utf-8") as archivo:
        return json.load(archivo)
