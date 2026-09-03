def en_rango(actualizar_estado, inicio, fin):
    """Mapea un progreso local 0-100 al intervalo [inicio, fin] sin tocar la lógica."""
    if actualizar_estado is None:
        return None

    def _callback(texto, local=0):
        try:
            local = int(local or 0)
        except (TypeError, ValueError):
            local = 0
        local = max(0, min(100, local))
        valor = int(round(inicio + (local / 100.0) * (fin - inicio)))
        actualizar_estado(texto, valor)

    return _callback
