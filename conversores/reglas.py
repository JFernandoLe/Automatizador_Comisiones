def fila_vacia(fila):

    for valor in fila:

        if str(valor).strip():
            return False

    return True

def cumple_filtro(
    fila,
    filtro,
    mapa_columnas
):

    columna = filtro["columna"]

    if columna not in mapa_columnas:
        return False

    indice = mapa_columnas[columna]

    dato = str(fila[indice]).upper()

    valor = str(filtro["valor"]).upper()

    operador = filtro["operador"]

    if operador == "contiene":
        return valor in dato

    elif operador == "igual":
        return dato == valor

    elif operador == "diferente":
        return dato != valor

    return False

def procesar_fila(
    fila,
    config,
    mapa_columnas,
    estadisticas
):

    estadisticas["filas_leidas"] += 1

    # Eliminar filas vacías
    if (
    config["eliminar_filas_vacias"]
    and fila_vacia(fila)
    ):
        estadisticas["filas_eliminadas"] += 1
        return None

    # Filtros
    for filtro in config["filtros"]:

        if cumple_filtro(
            fila,
            filtro,
            mapa_columnas
        ):
            estadisticas["filas_eliminadas"] += 1
            return None

    # Reglas de negocio
    fila = aplicar_reglas(
        fila,
        config["reglas"],
        mapa_columnas,
        estadisticas
    )

    return fila

def aplicar_reglas(
    fila,
    reglas,
    mapa_columnas,
    estadisticas
):

    for regla in reglas:

        if regla["tipo"] != "multiplicar":
            continue

        columna_condicion = regla["condicion"]["columna"]

        if columna_condicion not in mapa_columnas:
            continue

        columna_destino = regla["columna"]

        if columna_destino not in mapa_columnas:
            continue

        idx_condicion = mapa_columnas[columna_condicion]
        idx_destino = mapa_columnas[columna_destino]

        texto = str(
            fila[idx_condicion]
        ).upper()

        buscar = str(
            regla["condicion"]["valor"]
        ).upper()

        if buscar in texto:

            try:

                fila[idx_destino] = (
                    float(fila[idx_destino])
                    * regla["factor"]
                )
                estadisticas["filas_modificadas"] += 1

            except:
                pass

    return fila
