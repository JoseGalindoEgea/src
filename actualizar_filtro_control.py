import os
from pydoc import stripid
import shutil
import pandas as pd
import openpyxl
#from openpyxl import load_workbook
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
            df = pd.read_excel(origen, sheet_name=Mypestaña)
            datos[Mypestaña] = df
        return datos
    except FileNotFoundError:
        print(f"El archivo {origen} no se encontró.")
    except IOError:
        print(f"No se pudo leer el archivo {origen}.")

#Funcion: escribir_datos_en_excel(destino, datos)
def escribir_datos_en_excel(destino, datos):
    try:
        with pd.ExcelWriter(destino, engine='openpyxl') as writer:
            writer.book = load_workbook(destino)
            for Mypestaña, df in datos.items():
                df.to_excel(writer, sheet_name=Mypestaña, index=False)
            writer.save()
            writer.close()
    except FileNotFoundError:
        print(f"El archivo {destino} no se encontró.")
    except IOError:
        print(f"No se pudo leer el archivo {destino}.")

#Funcion copiado hojas excel con backup
#   cp_Excel_bck(rOrig, fOrig, shOrig, rDest, fDest, shDest)
def cp_Excel_bck(rOrig, fOrig, shOrig, rDest, fDest, shDest): # noqa: E999
    stError = False
    resultado = ""
    fBackup = fDest + '_v1'
    #1 paso) Backup del fichero filtro destino
    try:
        src_file = os.path.join(rDest, fDest)   #Fichero origen
        dst_file = os.path.join(rDest, fBackup)    #Fichero destino
        if not os.path.exists(src_file):
            print("Error, los ficheros origen o destino, no existen")
            stError = True
            resultado = "Error el fichero origen o destino no existen"
        else: 
            shutil.copy(src_file, dst_file)
            resultado = "Correcto, Backup archivo destino realizado"
            print(resultado + ' ' + fDest)
            stError = False
    except Exception as e:
        print(f"How exceptional! {e}")
        print("Error -Step1, no se pudo copiar el archivo de backup")
        stError = True
    

    #2 paso) Obtener el fichero origen

    if fOrig is not nullcontext and shOrig is not nullcontext and not stError:
        print("Fichero origen: ", fOrig)
        src_file = os.path.join(rOrig, fOrig)
        if not os.path.exists(src_file):
            print("Error, el fichero origen no existe", src_file)
            resultado = "Error el fichero origen no existe"
            stError = True
        else:
            # Load Excel File and give path to your file
            try:
                datosLec = cargar_datos_desde_excel(src_file, shOrig)
                resultado = "Correcto, Carga archivo origen"
                stError = False
                print(resultado + ' ' + fOrig)
            except FileNotFoundError:
                print("¡Error! No se encontró el archivo de origen o destino.")    
            except Exception as e:
                print(f"¡Error! Ocurrió un problema: {e}")
                print("Error, no se pudo abrir el fichero origen", src_file)
                resultado = "Error carga archivo origen"
                stError = True
               
    else:
        resultado = "Nombre vacio fichero origen"
        stError = True

    #3 paso) Escritura del df en un archivo excel
    if not stError:
        try:
            dst_file = os.path.join(rDest, fDest)
            datosEsc[shDest] = datosLec[shOrig]
            escribir_datos_en_excel(dst_file, datosEsc)
            print("Datos copiados con éxito en el archivo destino.")
            #df_Filtro.to_excel(dst_file, sheet_name = shDest, index=False)
            resultado = "Correcto, Escribe archivo destino"
            stError = False
            print(resultado + ' ' + dst_file)
        except FileNotFoundError:
            print("¡Error! No se encontró el archivo de origen o destino.")    
        except Exception as e:
            print(f"¡Error! Ocurrió un problema: {e}")
            print("Error, no se pudo grabar el archivo destino")
            resultado = "Error escritura archivo destino"
            stError = True

    #4 paso) Resultado
        return stError




def main():
    ruta_parametros = "D:\\IBD.GIT\\Python\\Automatizaciones\\src\\"
    archivo_parametros = "actualizar_filtro_parametros.txt"
    fparametros = ruta_parametros + archivo_parametros

    parametros = leer_parametros(fparametros)

    #Parametros leidos de fichero
    for ruta in parametros:
        if "RUTA_ARCHIVO_DESCARGAS" == ruta:
            ruta_archivo_descargas = str(parametros["RUTA_ARCHIVO_DESCARGAS"])
        if "RUTA_ARCHIVO_FILTROS" == ruta:
            ruta_archivo_filtros = str(parametros["RUTA_ARCHIVO_FILTROS"])
        if "ARCHIVO_FILTRO" == ruta:
            nombre_dst_filtro = str(parametros["ARCHIVO_FILTRO"])
        if "ARCHIVO_TRELLO" == ruta:
            nombre_dst_trello = str(parametros["ARCHIVO_TRELLO"])
        if "PEST_FILTRO_ORIGEN" == ruta:
            pest_src_filtro = str(parametros["PEST_FILTRO_ORIGEN"])
        if "PEST_FILTRO_DESTINO" == ruta:
            pest_dst_filtro = str(parametros["PEST_FILTRO_DESTINO"])
        if "PEST_TRELLO_ORIGEN" == ruta:
            pest_src_trello = str(parametros["PEST_TRELLO_ORIGEN"])
        if "PEST_TRELLO_DESTINO" == ruta:
            pest_dst_trello = str(parametros["PEST_TRELLO_DESTINO"])

    #1a ejecución con Filtro
    fNombreOrigen = str(input("Introduzca nombre fichero origen (filtro): "))
    fSrcOrigen = ruta_archivo_descargas + fNombreOrigen + ".xlsx"
    src_hoja = pest_src_filtro.strip()
    dst_hoja = pest_dst_filtro.strip()
    stResultadoErr = False
    stResultadoErr = cp_Excel_bck(rOrig = ruta_archivo_descargas, fOrig = fSrcOrigen, 
                            shOrig = src_hoja, rDest = ruta_archivo_filtros, 
                            fDest = nombre_dst_filtro, shDest = dst_hoja)
    print(f'Resultado de copiar_planillas_backup: stResultadoErr = {stResultadoErr}')
        
    stContinuar = str(input("Desea continuar (S/N): "))
    if (stContinuar.strip() == "S") or (stContinuar.strip() == "s"):

        fNombreOrigen = str(input("Introduzca nombre fichero origen (trello): "))
        fSrcOrigen = ruta_archivo_descargas + fNombreOrigen + ".xlsx"
        src_hoja = pest_src_trello.strip()
        dst_hoja = pest_dst_trello.strip()

        stResultadoErr = False
        stResultadoErr = cp_Excel_bck(rOrig = ruta_archivo_descargas, fOrig = fSrcOrigen, 
                            shOrig = src_hoja, rDest = ruta_archivo_filtros, 
                            fDest = nombre_dst_trello, shDest = dst_hoja)
        print(f'Resultado de copiar_planillas_backup (segundo): stResultadoErrs = {stResultadoErr}')
    else:
        print("Finalizado.")


if __name__ == "__main__":
    main()
