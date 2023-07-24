import os
from pydoc import stripid
import shutil
import pandas as pd
import openpyxl
#from openpyxl import load_workbook
from contextlib import nullcontext



#Funcion copiado hojas excel con backup
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
    
    #1.bis. Fichero origen
    print("Fichero origen: ", fOrig)
    src_file = os.path.join(rOrig, fOrig)
    if not os.path.exists(src_file):
        print("Error, el fichero origen no existe", src_file)
        stError = True
    else:
        fBackup = "Filtro_origen_backup_v1"
        dst_file = os.path.join(rOrig, fBackup)
        try:
            shutil.copy(src_file, dst_file)
            resultado = "Correcto, Backup origen realizado"
            print(resultado + ' ' + fOrig)
            stError = False
        except Exception as e:
            print(f"How exceptional! {e}")
            print("Error copia fichero origen")
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
                df_Filtro = pd.read_excel(src_file, sheet_name= shOrig)
                resultado = "Correcto, Carga archivo origen"
                stError = False
                print(resultado + ' ' + fOrig)
                
            except Exception as e:
                print(f"How exceptional! {e}")
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
    
            df_Filtro.to_excel(dst_file, sheet_name = shDest, index=False)
            resultado = "Correcto, Escribe archivo destino"
            stError = False
            print(resultado + ' ' + dst_file)
        except Exception as e:
            print(f"How exceptional! {e}")
            print("Error, no se pudo grabar el archivo destino")
            resultado = "Error escritura archivo destino"
            stError = True

    #4 paso) Resultado
        return stError


#Proceso principal
# Give the location of the file
My_downloads = "C:\\Users\\magaegjf\\Downloads\\"
My_Tron_path = "C:\\Users\\magaegjf\\OneDrive - AYESA\\SVC-MAPFRE-TRON\\01.GESTION\\1.PRODUCCION\\2.Avances\\CdM"
fTarget_file = "TRON.AM.IBERMATICA.FULLANNUAL.xlsx"
fTarget_backup ="TRON.AM.IBERMATICA.FULLANNUAL_v1.xlsx"
fNombreOrigen = str(input("Introduzca nombre fichero origen (filtro): "))
fSrcOrigen = My_downloads + fNombreOrigen + ".xlsx"
src_hoja = "Your Jira Issues"
dst_hoja = "TRON.FULLANNUAL"

stResultadoErr = False
stResultadoErr = cp_Excel_bck(rOrig = My_downloads, fOrig = fSrcOrigen, 
                            shOrig = src_hoja, rDest = My_Tron_path, 
                            fDest = fTarget_file, shDest = dst_hoja)
print(f'Resultado de copiar_planillas_backup: stResultadoErr = {stResultadoErr}')

stContinuar = str(input("Desea continuar (S/N): "))
if (stContinuar.strip() == "S") or (stContinuar.strip() == "s"):
    fTarget_file = "Trello-Mantenimiento TRON Ibermatica_carga.xlsx"
    fTarget_backup ="Trello-Mantenimiento TRON Ibermatica_carga_v1.xlsx"
    fNombreOrigen = str(input("Introduzca nombre fichero origen (trello): "))
    fSrcOrigen = My_downloads + fNombreOrigen + ".xlsx"
    src_hoja = "Trello Export"
    dst_hoja = "Carga Trello"

    stResultadoErr = False
    stResultadoErr = cp_Excel_bck(rOrig = My_downloads, fOrig = fSrcOrigen, 
                            shOrig = src_hoja, rDest = My_Tron_path, 
                            fDest = fTarget_file, shDest = dst_hoja)
    print(f'Resultado de copiar_planillas_backup (segundo): stResultadoErrs = {stResultadoErr}')
else:
    print("Finalizado.")