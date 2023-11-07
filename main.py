import pandas as pd
import time
from function.funciones import crear_df, carga_archivos, seleccion, concatenar, crearhoja


# Mensaje de inicio
print("Bienvenidos al constructor de la planilla del relevamiento energetico")
# Seleccionar el archivo a procesar
print("Seleccionar el archivo a utilizar")
time.sleep(4)


# Unir columnas y filas por seccion
# carga y elaboración de df
archivo = carga_archivos()
df = crear_df(archivo)

# calculo de registros de una variable
numero_de_filas = df.shape[0]
print(numero_de_filas)

# Seleccionar filas y columnas del Excel

# seleccionar nombres de los espacios
sel_nombre = seleccion(df, 0, numero_de_filas, 1, 7)

# seleccionar detalles de los espacios
sel_detalle_espacio = seleccion(df, 0, numero_de_filas, 8, 13)

# seleccionar climatizacion
sel_clima_1 = seleccion(df, 0, numero_de_filas, 18, 35)
# Tomar nombres de columnas de la seleccion 1
nombre_col_clima = sel_clima_1.columns
sel_clima_2 = seleccion(df, 0, numero_de_filas, 37, 54)
# Remplazar nombres de columnas de la seleccion 2 con los datos de la columna 1
sel_clima_2.columns = nombre_col_clima
sel_clima_3 = seleccion(df, 0, numero_de_filas, 56, 73)
# Remplazar nombres de columnas de la seleccion 3 con los datos de la columna 1
sel_clima_3.columns = nombre_col_clima


# seleccionar iluminacion
sel_ilu_1 = seleccion(df, 0, numero_de_filas, 75, 83)
nombre_col_ilu = sel_ilu_1.columns
sel_ilu_2 = seleccion(df, 0, numero_de_filas, 85, 93)
sel_ilu_2.columns = nombre_col_ilu
sel_ilu_3 = seleccion(df, 0, numero_de_filas, 95, 103)
sel_ilu_3.columns = nombre_col_ilu


# seleccionar Equipamiento de oficina/informatico
sel_info_1 = seleccion(df, 0, numero_de_filas, 105, 114)
# elimina numeros y puntos del nombre de las columnas
sel_info_1.columns = sel_info_1.columns.str.replace(r'\.\d+', '', regex=True)
nombre_col_info = sel_info_1.columns
sel_info_2 = seleccion(df, 0, numero_de_filas, 116, 125)
sel_info_2.columns = nombre_col_info
sel_info_3 = seleccion(df, 0, numero_de_filas, 127, 136)
sel_info_3.columns = nombre_col_info
sel_info_4 = seleccion(df, 0, numero_de_filas, 138, 147)
sel_info_4.columns = nombre_col_info
sel_info_5 = seleccion(df, 0, numero_de_filas, 149, 158)
sel_info_5.columns = nombre_col_info

# Seleccionar Equipamiento de industrial, taller, laboratorio
sel_indu_1 = seleccion(df, 0, numero_de_filas, 160, 170)
# elimina numeros y puntos del nombre de las columnas
sel_indu_1.columns = sel_indu_1.columns.str.replace(r'\.\d+', '', regex=True)
nombre_col_ind = sel_indu_1.columns
sel_indu_2 = seleccion(df, 0, numero_de_filas, 172, 182)
sel_indu_2.columns = sel_indu_1.columns
sel_indu_3 = seleccion(df, 0, numero_de_filas, 184, 194)
sel_indu_3.columns = sel_indu_1.columns
sel_indu_4 = seleccion(df, 0, numero_de_filas, 196, 206)
sel_indu_4.columns = sel_indu_1.columns
sel_indu_5 = seleccion(df, 0, numero_de_filas, 208, 218)
sel_indu_5.columns = sel_indu_1.columns

# seleccionar electrodomesticos
sel_electro_1 = seleccion(df, 0, numero_de_filas, 220, 228)
sel_electro_1.columns = sel_electro_1.columns.str.replace(
    r'\.\d+', '', regex=True)
nombre_columns_electro = sel_electro_1.columns
sel_electro_2 = seleccion(df, 0, numero_de_filas, 230, 238)
sel_electro_2.columns = nombre_columns_electro
sel_electro_3 = seleccion(df, 0, numero_de_filas, 240, 248)
sel_electro_3.columns = nombre_columns_electro
sel_electro_4 = seleccion(df, 0, numero_de_filas, 250, 258)
sel_electro_4.columns = nombre_columns_electro
sel_electro_5 = seleccion(df, 0, numero_de_filas, 260, 268)
sel_electro_5.columns = nombre_columns_electro

# Union de nombre con datos
UnirConDatos = concatenar(sel_nombre, sel_detalle_espacio, 1, None)

# Nombre y climatizacion
# se concatena horizontalmente el nombre con cada seleccion de columnas
unir_clima_1 = concatenar(sel_nombre, sel_clima_1, 1, None)
unir_clima_2 = concatenar(sel_nombre, sel_clima_2, 1, None)
unir_clima_3 = concatenar(sel_nombre, sel_clima_3, 1, None)
# se concatenan verticalmente cada conjunto de nombre con seleccion para que solo quede una columna de informacion
unir_clima_v = pd.concat([unir_clima_1, unir_clima_2, unir_clima_3], axis=0)
unir_clima_v.dropna(subset='Tipo de distribucion', inplace=True)

# Nombre y iluminacion (misma logica que climatizacion)
unir_ilu_1 = concatenar(sel_nombre, sel_ilu_1, 1, False)
unir_ilu_2 = concatenar(sel_nombre, sel_ilu_2, 1, False)
unir_ilu_3 = concatenar(sel_nombre, sel_ilu_3, 1, False)
unir_ilu_v = pd.concat([unir_ilu_1, unir_ilu_2, unir_ilu_3], axis=0)
unir_ilu_v.dropna(subset='Tecnología de iluminación', inplace=True)

# Nombre y Oficina/informatico
unir_info_1 = concatenar(sel_nombre, sel_info_1, 1, False)
unir_info_2 = concatenar(sel_nombre, sel_info_2, 1, False)
unir_info_3 = concatenar(sel_nombre, sel_info_3, 1, False)
unir_info_4 = concatenar(sel_nombre, sel_info_4, 1, False)
unir_info_5 = concatenar(sel_nombre, sel_info_5, 1, False)
unir_info_v = pd.concat(
    [unir_info_1, unir_info_2, unir_info_3, unir_info_4, unir_info_5], axis=0)
unir_info_v.dropna(subset='Tipo de Equipamiento', inplace=True)
# unirInformaticoVertical.dropna(subset=)

# Nombre y industrial/laboratorio
unir_indu_1 = concatenar(sel_nombre, sel_indu_1, 1, False)
unir_indu_2 = concatenar(sel_nombre, sel_indu_2, 1, False)
unir_indu_3 = concatenar(sel_nombre, sel_indu_3, 1, False)
unir_indu_4 = concatenar(sel_nombre, sel_indu_4, 1, False)
unir_indu_5 = concatenar(sel_nombre, sel_indu_5, 1, False)
unir_indu_v = pd.concat(
    [unir_indu_1, unir_indu_2, unir_indu_3, unir_indu_4, unir_indu_5], axis=0)
unir_indu_v.dropna(subset='Tamaño', inplace=True)

# Nombre y electrodomestico
unir_electro_1 = concatenar(sel_nombre, sel_electro_1, 1, False)
unir_electro_2 = concatenar(sel_nombre, sel_electro_2, 1, False)
unir_electro_3 = concatenar(sel_nombre, sel_electro_3, 1, False)
unir_electro_4 = concatenar(sel_nombre, sel_electro_4, 1, False)
unir_electro_5 = concatenar(sel_nombre, sel_electro_5, 1, False)
unir_electro_v = pd.concat([unir_electro_1, unir_electro_2,
                           unir_electro_3, unir_electro_4, unir_electro_5], axis=0)
unir_electro_v.dropna(subset='Tipo de electrodomestico', inplace=True)


# hoja de excel de salida

path2 = 'datosretorno/tabla.xlsx'

# armado de hojas por seccion
armarExcel = pd.ExcelWriter(path2, engine='xlsxwriter')

# hoja de datos
crearhoja(UnirConDatos, armarExcel, 'DetalleEspacio')
# hoja de climatizacion
crearhoja(unir_clima_v, armarExcel, 'Climatizacion')
# hoja de iluminacion
crearhoja(unir_ilu_v, armarExcel, 'Iluminación')
# hoja oficina/informatica
crearhoja(unir_info_v, armarExcel, 'Oficina.Informatico')
# hoja industrial/laboratorio
crearhoja(unir_indu_v, armarExcel, 'Industria.Laboratorio')
# hoja electrodomesticos
crearhoja(unir_electro_v, armarExcel, 'Electrodomesticos')

armarExcel.close()

print('El escaneo a finalizado correctamente')