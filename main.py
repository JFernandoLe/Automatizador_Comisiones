import os
import sys
from pathlib import Path


def _preparar_entorno():
    if getattr(sys, "frozen", False):
        os.chdir(Path(sys.executable).resolve().parent)
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "MetLife.CommiFlow"
            )
        except Exception:
            pass


_preparar_entorno()

from gui.ventana_principal import iniciar_app

if __name__ == "__main__":
    iniciar_app()
