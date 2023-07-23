iimport os
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
My_path = "D:\\IBD.GIT\\Python\\Automatizaciones\\demo\\"
My_Tron_path = "C:\\Users\\magaegjf\\OneDrive - AYESA\\SVC-MAPFRE-TRON\\01.GESTION\\1.PRODUCCION\\2.Avances\\CdM"
fSource_file = "demo.xlsx"
fTarget_file = "demo1.xlsx"
fTarget_filter_file = "TRON.AM.IBERMATICA.FULLANNUAL.xlsx"
fTarget_filter_backup ="TRON.AM.IBERMATICA.FULLANNUAL_v1.xlsx"
fOrigen = My_path + fSource_file
fDestino = My_path + fTarget_file
resultado = ""

#datos_origen = pd.read_excel(fOrigen)
#sales_data = pd.read_excel(fOrigen)

#Backup del fichero filtro de origen
try:
    src_file = os.path.join(My_Tron_path, fTarget_filter_file)   #Fichero origen
    dst_file = os.path.join(My_Tron_path, fTarget_filter_backup) #Fichero destino
    shutil.copy(src_file, dst_file)
    resultado = "Backup filtro realizado"
except:
    print("Error, no se pudo copiar el archivo de backup")

# Parsing archivo lectura
#pd.read_excel("path_to_file.xls", "Sheet1", index_col=None, na_values=["NA"])

fNombreOrigen = str(input("Introduzca nombre fichero origen: "))
if fNombreOrigen is not nullcontext:
    fOrigen = My_downloads + fNombreOrigen + ".xlsx"
    # Load Excel File and give path to your file
    try:
        df_Filtro = pd.read_excel(fOrigen, sheet_name="Your Jira Issues")
        resultado = "Carga archivo origen"
    except:
        print("Error, no se pudo abrir el fichero origen")
        resultado = "Error carga archivo origen"
else:
    resultado = "Nombre vacio fichero origen"

# Escritura del df en un archivo excel
dst_file = os.path.join(My_Tron_path, fTarget_filter_file)
df_Filtro.to_excel(dst_file, sheet_name ='TRON.FULLANNUAL')


# Perform Data Manipulation
#df = df[df['Sales'] > 90]
#sales_data = sales_data.groupby('Product').sum()

# Write Data to Excel File
#book = load_workbook(fOrigen)
#writer = pd.ExcelWriter(fDestino, engine='openpyxl')
#writer.book = book
#df.to_excel(writer, index=False)
#writer.save()
#writer.close()

#Only with panda
#writer = pd.ExcelWriter(fDestino)
#sales_data.to_excel(writer, sheet_name ='TRON.FULLANNUAL')
#writer.save()
#writer.close()

