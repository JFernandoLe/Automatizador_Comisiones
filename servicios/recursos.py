import sys
from pathlib import Path


def ruta_base():
    if getattr(sys, "frozen", False) and getattr(sys, "_MEIPASS", None):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


def ruta_logo():
    base = ruta_base()
    for nombre in ("metlife_logo.jpg", "metlife_logo.png"):
        candidato = base / "assets" / nombre
        if candidato.exists():
            return candidato
    return None


def ruta_icono():
    base = ruta_base()
    for candidato in (base / "icono.ico", base.parent / "icono.ico"):
        if candidato.exists():
            return candidato
    return None
