import pandas as pd
from openpyxl import load_workbook

def cargar_datos_desde_excel(origen, pestañas):
    datos = {}
    for pestaña in pestañas:
        df = pd.read_excel(origen, sheet_name=pestaña)
        datos[pestaña] = df
    return datos

def escribir_datos_en_excel(destino, datos):
    with pd.ExcelWriter(destino, engine='openpyxl') as writer:
        writer.book = load_workbook(destino)
        for pestaña, df in datos.items():
            df.to_excel(writer, sheet_name=pestaña, index=False)
        writer.save()

def main():
    archivo_origen = input("Ingrese el nombre del archivo Excel origen: ")
    archivo_destino = input("Ingrese el nombre del archivo Excel destino: ")
    pestañas_a_copiar = ["pestaña1", "pestaña2", "pestaña3"]  # Agrega aquí las pestañas a copiar desde el archivo origen

    try:
        datos = cargar_datos_desde_excel(archivo_origen, pestañas_a_copiar)
        escribir_datos_en_excel(archivo_destino, datos)
        print("Datos copiados con éxito en el archivo destino.")
    except FileNotFoundError:
        print("¡Error! No se encontró el archivo de origen o destino.")
    except Exception as e:
        print(f"¡Error! Ocurrió un problema: {e}")

    #pestañas_a_copiar = []
    #while True:
    #    pestaña = input("Ingrese el nombre de una pestaña a copiar (o escriba 'fin' para terminar): ")
    #    if pestaña.lower() == "fin":
    #        break
    #    pestañas_a_copiar.append(pestaña)

    #try:
    #    datos_a_copiar = leer_datos_desde_excel(archivo_excel_origen, pestañas_a_copiar)
    #    escribir_datos_en_excel(datos_a_copiar, archivo_excel_destino)
    #    print("Datos copiados exitosamente al archivo Excel destino.")
    #except Exception as e:
    #    print(f"Error: {e}")

if __name__ == "__main__":
    main()
