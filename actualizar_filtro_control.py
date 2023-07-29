import os
from pydoc import stripid
import shutil
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from contextlib import nullcontext

#Funcion: leer_parametros(archivo_parametros)
def leer_parametros(archivo_parametros):
    try:
        parametros = {}
        with open(archivo_parametros, "r") as archivo:
            for linea in archivo:
                etiqueta, valor = linea.strip().split("|")
                parametros[etiqueta] = valor
        return parametros
    except FileNotFoundError:
        print(f"Error, El archivo {archivo_parametros} no se encontró.")
    except IOError:
        print(f"Error, No se pudo leer el archivo {archivo_parametros}.")


#Funcion: cargar_datos_desde_excel(origen, pestañas)
def cargar_datos_desde_excel(origen, pestañas):
    try:
        datos = {}
        for Mypestaña in pestañas:
            #Lectura archivo excel según pestaña
            df = pd.read_excel(origen, sheet_name=Mypestaña)
            datos[Mypestaña] = df
        return datos
    except FileNotFoundError:
        print(f"Error, El archivo {origen} no se encontró.")
    except IOError:
        print(f"Error, No se pudo leer el archivo {origen}.")

#Funcion: escribir_datos_en_excel(destino, datos)
#Informacion: datos es un array por pestaña
def escribir_datos_en_excel(destino, datos):
    try:
        Mypestaña = ""
        with pd.ExcelWriter(destino, engine='openpyxl') as writer:
            writer.book = load_workbook(destino)
            for Mypestaña, df in datos.items():
                df.to_excel(writer, sheet_name=Mypestaña, index=False)
            writer.save()
            writer.close()
    except FileNotFoundError:
        print(f"Error, El archivo {destino} no se encontró.")
    except IOError:
        print(f"Error, No se pudo leer el archivo {destino} ni pestaña {Mypestaña}.")

def archive_backup(src_file, dst_file):
    stError = False
    try:
        if not os.path.exists(src_file):
            print(f"Error, el fichero origen, no existe: {src_file}")
            stError = True
        else: 
            shutil.copy(src_file, dst_file)
            print("Backup realizado.")
            stError = False
    except Exception as e:
        print(f"How exceptional! {e}")
        print(f"Error, no se pudo realizar el backup del fichero destino")
        stError = True

    return stError


#Funcion copiado hojas excel con backup
#   cp_Excel_bck(rOrig, fOrig, shOrig, rDest, fDest, shDest)
def cp_Excel_bck( mi_parametros ): # noqa: E999
    stError = False
    #descarga del diccionario de parmetros
    rOrig = mi_parametros["rOrig"]
    fOrig = mi_parametros["fOrig"]
    shOrig = mi_parametros["shOrig"]
    rDest = mi_parametros["rDest"]
    fDest = mi_parametros["fDest"]
    shDest = mi_parametros["shDest"]

    resultado = ""
    fBackup = fDest + '_v1'
    #1 paso) Backup del fichero filtro destino
    if rDest is not nullcontext and fDest is not nullcontext:
        src_file = os.path.join(rDest, fDest)   #Fichero origen
        dst_file = os.path.join(rDest, fBackup)    #Fichero destino
        
        #Backup archivo origen
        stError = archive_backup(src_file, dst_file)
    else:
        stError = True
   
    #2 paso) Obtener el fichero origen
    if fOrig is not nullcontext and rOrig is not nullcontext and shOrig is not nullcontext and not stError:
        print("Step-2, Fichero origen: ", fOrig)
        src_file = os.path.join(rOrig, fOrig)
        if not os.path.exists(src_file):
            print(f"Error, el fichero origen no existe: {src_file}")
            resultado = "Error, el fichero origen no existe"
            stError = True
        else:
            # Load Excel File and give path to your file
            datosLec = {}
            datosLec = cargar_datos_desde_excel(src_file, shOrig)
            resultado = "Correcto, Carga archivo origen"
            stError = False
            print(resultado + ' ' + fOrig)

               
    else:
        resultado = "Error, Nombre vacio fichero origen/ruta origen/pestaña origen"
        stError = True

    #3 paso) Escritura del df en un archivo excel
    if fDest is not nullcontext and shDest is not nullcontext and not stError:
        datosEsc = {}
        dst_file = os.path.join(rDest, fDest)
        datosEsc[shDest] = datosLec[shOrig]
        escribir_datos_en_excel(dst_file, datosEsc)
        print(f"Datos copiados con éxito en el archivo destino: {dst_file}")
        #df_Filtro.to_excel(dst_file, sheet_name = shDest, index=False)
        resultado = "Correcto, Escribe archivo destino"
        stError = False

    else:
        resultado = "Incorrecto, hubo algun problema"
        print(resultado + ' ' + dst_file)
        stError = True
        

    #4 paso) Resultado
        return stError




def main():
    ruta_parametros = "D:\\IBD.GIT\\Python\\Automatizaciones\\src\\"
    archivo_parametros = "actualizar_filtro_parametros.txt"
    fparametros = ruta_parametros + archivo_parametros

    parametros = leer_parametros(fparametros)

    # Paso 1. Carga parametros fichero
    #Parametros leidos de fichero
    for ruta in parametros:
        if "RUTA_ARCHIVO_DESCARGAS" == ruta:
            ruta_archivo_descargas = str(parametros["RUTA_ARCHIVO_DESCARGAS"]).strip()
        if "RUTA_ARCHIVO_FILTROS" == ruta:
            ruta_archivo_filtros = str(parametros["RUTA_ARCHIVO_FILTROS"]).strip()
        if "ARCHIVO_FILTRO" == ruta:
            nombre_dst_filtro = str(parametros["ARCHIVO_FILTRO"]).strip()
        if "ARCHIVO_TRELLO" == ruta:
            nombre_dst_trello = str(parametros["ARCHIVO_TRELLO"]).strip()
        if "PEST_FILTRO_ORIGEN" == ruta:
            pest_src_filtro = str(parametros["PEST_FILTRO_ORIGEN"]).strip()
        if "PEST_FILTRO_DESTINO" == ruta:
            pest_dst_filtro = str(parametros["PEST_FILTRO_DESTINO"]).strip()
        if "PEST_TRELLO_ORIGEN" == ruta:
            pest_src_trello = str(parametros["PEST_TRELLO_ORIGEN"]).strip()
        if "PEST_TRELLO_DESTINO" == ruta:
            pest_dst_trello = str(parametros["PEST_TRELLO_DESTINO"]).strip()

    # Paso 2.- Preparar la gestión del filtro
    #1a ejecución con Filtro
    fNombreOrigen = str(input("Introduzca nombre fichero origen (filtro): "))
    fSrcOrigen = ruta_archivo_descargas + fNombreOrigen + ".xlsx"
    src_hoja = pest_src_filtro.strip()
    dst_hoja = pest_dst_filtro.strip()
    stResultadoErr = False

    #lista de valores diccionario
    prm_cp_bck = {
        "rOrig": [],
        "fOrig": [],
        "shOrig": [],
        "rDest": [],
        "fDest": [],
        "shDest": [],
    }

    #Informar el diccionario
    prm_cp_bck["rOrig"] = ruta_archivo_descargas    #parametro del fichero
    prm_cp_bck["fOrig"] = fSrcOrigen
    prm_cp_bck["shOrig"] = src_hoja
    prm_cp_bck["rDest"] = ruta_archivo_filtros      #parametro del fichero
    prm_cp_bck["fDest"] = nombre_dst_filtro         #parametro del fichero
    prm_cp_bck["shDest"] = dst_hoja

    #prm_cp_bck["rOrig"].append(ruta_archivo_descargas)
    #prm_cp_bck["fOrig"].append(fSrcOrigen)
    #prm_cp_bck["shOrig"].append(src_hoja)
    #prm_cp_bck["rDest"].append(ruta_archivo_filtros)
    #prm_cp_bck["fDest"].append(nombre_dst_filtro)
    #prm_cp_bck["shDest"].append(dst_hoja)
     

    for llave in prm_cp_bck:
        print("Llamada filtro: ", llave, ": ", prm_cp_bck[llave])

    #print(f"Elementos variables: ruta descargas: {ruta_archivo_descargas}; fOrig: {fSrcOrigen}; shOrig: {src_hoja}; rDest: {ruta_archivo_filtros}; fDest: {nombre_dst_filtro}; shDest: {dst_hoja}")
    #print("")
    
    
    #stResultadoErr = cp_Excel_bck(rOrig = ruta_archivo_descargas, fOrig = fSrcOrigen, 
    #                        shOrig = src_hoja, rDest = ruta_archivo_filtros, 
    #                        fDest = nombre_dst_filtro, shDest = dst_hoja)
    stResultadoErr = cp_Excel_bck( prm_cp_bck)
    print(f'Resultado de copiar_planillas_backup: stResultadoErr = {stResultadoErr}')

    #Paso 3.- Preparar la gestión del fichero Trello    
    stContinuar = str(input("Desea continuar (S/N): "))
    if (stContinuar.strip() == "S") or (stContinuar.strip() == "s"):

        fNombreOrigenTrello = str(input("Introduzca nombre fichero origen (trello): "))
        fSrcOrigen = ruta_archivo_descargas + fNombreOrigenTrello + ".xlsx"
        src_hoja = pest_src_trello.strip()
        dst_hoja = pest_dst_trello.strip()

        #Informar el diccionario de parametros
        prm_cp_bck["rOrig"] = ruta_archivo_descargas    #parametro del fichero
        prm_cp_bck["fOrig"] = fSrcOrigen
        prm_cp_bck["shOrig"] = src_hoja
        prm_cp_bck["rDest"] = ruta_archivo_filtros      #parametro del fichero
        prm_cp_bck["fDest"] = nombre_dst_trello         #parametro del fichero
        prm_cp_bck["shDest"] = dst_hoja

        for llave in prm_cp_bck:
            print("Llamada Trello: ", llave, ": ", prm_cp_bck[llave])

        stResultadoErr = False
        stResultadoErr = cp_Excel_bck(prm_cp_bck)
        print(f'Resultado de copiar_planillas_backup (segundo): stResultadoErrs = {stResultadoErr}')
    else:
        print("Finalizado.")


if __name__ == "__main__":
     # Only run the main function if this module is being run directly with `python main.py` or `python -m main`
    main()
