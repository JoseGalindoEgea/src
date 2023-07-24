def leer_parametros(archivo_parametros):
    parametros = {}
    with open(archivo_parametros, "r") as archivo:
        for linea in archivo:
            etiqueta, valor = linea.strip().split(":")
            parametros[etiqueta] = valor
    return parametros

def procesar_archivo(ruta_archivo):
    try:
        with open(ruta_archivo, "r") as archivo:
            contenido = archivo.read()
            print("Contenido del archivo:")
            print(contenido)
    except FileNotFoundError:
        print(f"El archivo {ruta_archivo} no se encontró.")
    except IOError:
        print(f"No se pudo leer el archivo {ruta_archivo}.")

def main():
    
    archivo_parametros = "actualizar_filtro_parametros.txt"
    parametros = leer_parametros(archivo_parametros)

    for ruta in parametros:
        if "RUTA_ARCHIVO_DESCARGAS" == ruta:
            ruta_archivo = parametros["RUTA_ARCHIVO_DESCARGAS"]

        if "RUTA_ARCHIVO_FILTROS" == ruta:
            ruta_archivo = parametros["RUTA_ARCHIVO_FILTROS"]
        

    if "RUTA_ARCHIVO_1" in parametros:
        ruta_archivo = parametros["RUTA_ARCHIVO_1"]
        procesar_archivo(ruta_archivo)
    else:
        print("La etiqueta 'RUTA_ARCHIVO_1' no se encontró en el archivo de parámetros.")

if __name__ == "__main__":
    main()
