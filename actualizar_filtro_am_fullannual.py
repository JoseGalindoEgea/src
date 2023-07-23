import os
import shutil
import pandas as pd
from openpyxl import load_workbook
from contextlib import nullcontext

# General terminology translation
#Dataframe-worksheet
#series-column
#index-row headings
#row-row
#NaN - empty cell

# Give the location of the file
My_downloads = "C:\\Users\\magaegjf\\Downloads\\"
My_Tron_path = "C:\\Users\\magaegjf\\OneDrive - AYESA\\SVC-MAPFRE-TRON\\01.GESTION\\1.PRODUCCION\\2.Avances\\CdM"
fTarget_filter_file = "TRON.AM.IBERMATICA.FULLANNUAL.xlsx"
fTarget_filter_backup ="TRON.AM.IBERMATICA.FULLANNUAL_v1.xlsx"
resultado = ""
stError = False

#datos_origen = pd.read_excel(fOrigen)
#sales_data = pd.read_excel(fOrigen)

#Backup del fichero filtro de origen
try:
    src_file = os.path.join(My_Tron_path, fTarget_filter_file)   #Fichero origen
    dst_file = os.path.join(My_Tron_path, fTarget_filter_backup) #Fichero destino
    shutil.copy(src_file, dst_file)
    resultado = "Backup filtro realizado"
    print(resultado)
    stError = False
except:
    print("Error, no se pudo copiar el archivo de backup")
    stError = True

# Parsing archivo lectura
#pd.read_excel("path_to_file.xls", "Sheet1", index_col=None, na_values=["NA"])

fNombreOrigen = str(input("Introduzca nombre fichero origen: "))
if fNombreOrigen is not nullcontext:
    fOrigen = My_downloads + fNombreOrigen + ".xlsx"
    print("Fichero origen: ", fOrigen)
    if not os.path.exists(fOrigen):
        print("Error, el fichero origen no existe")
        resultado = "Error el fichero origen no existe"
        stError = True
    else:
        # Load Excel File and give path to your file
        try:
            src_hoja = "Your Jira Issues"
            df_Filtro = pd.read_excel(fOrigen, sheet_name= src_hoja)
            resultado = "Carga archivo origen"
            stError = False
            print(resultado + ' ' + fOrigen)
            m_row = df_Filtro.max_row
            print("Filas leidas: ", m_row)
        except:
            print("Error, no se pudo abrir el fichero origen")
            resultado = "Error carga archivo origen"
            stError = True
else:
    resultado = "Nombre vacio fichero origen"
    stError = True

# Escritura del df en un archivo excel
try:
    dst_file = os.path.join(My_Tron_path, fTarget_filter_file)
    dst_hoja = "TRON.FULLANNUAL"
    df_Filtro.to_excel(dst_file, sheet_name = dst_hoja, index=False)
    resultado = "Escribe archivo destino"
    stError = False
    print(resultado + ' ' + dst_file)
except:
    print("Error, no se pudo grabar el archivo destino")
    resultado = "Error escritura archivo destino"
    stError = True


