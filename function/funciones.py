import tkinter as tk
import pandas as pd
from tkinter import filedialog


# funciones de carga y creacion del data frame

# crear el data frame
def crear_df(path):
    datos = pd.read_excel(path)
    df1 = pd.DataFrame(datos)
    return df1

# Carga de archivo (hoja de calculo)
def carga_archivos():
    root = tk.Tk()
    root.withdraw()
    # global listadearchivos
    listadearchivos = filedialog.askopenfilenames(
        filetypes=[("Archivos de hojas de cálculo XLS", "*.xls *.xlsx *.ods *.csv")])
    print(listadearchivos)
    listadearchivos = str(listadearchivos[0])
    tipo = type(listadearchivos)
    print(tipo)
    return listadearchivos
    # print("Ubicación del archivo PDF seleccionado:", pdf_path)

# # Crear dataframe a partir de archivo
# def dfdearchivo():
#     archivo = carga_archivos()
#     df = crear_df(archivo)
#     return df

# def contadordefilas(dataframe):    
#     # calculo de registros de una variable
#     numeroDeFilas = df.shape[0]
#     return numeroDeFilas

def seleccion(df, unoregistro, totalfilas, iniciocolumnas, fincolumnas):
    return df.iloc[unoregistro:totalfilas, iniciocolumnas:fincolumnas]

# Funcion concat


def concatenar(dato1, dato2, orientacion, index):
    return pd.concat([dato1, dato2], axis=orientacion, ignore_index=index)

def crearhoja(variable, armarexcel, nombrehoja):
    variable.to_excel(armarexcel, sheet_name=nombrehoja, index=False)