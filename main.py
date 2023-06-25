import pandas as pd
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import locale

def clear_console():
    # Función para borrar la consola
    os.system('cls' if os.name == 'nt' else 'clear')

def filter_dataframe(df, llaves):
    # Eliminar las primeras 23 filas
    df = df.iloc[23:]

    # Eliminar la primera y última columna
    df = df.iloc[:, 1:-1]

    # Asignar la primera fila como nombres de columna
    df.columns = df.iloc[0]

    # Eliminar la primera fila, ya que ahora son los nombres de columna
    df = df.iloc[1:]

    # Filtrar filas con valores no nulos en la columna 'Canal o Sucursal'
    df_filtrado = df.dropna(subset=['Canal o Sucursal'])

    # Filtrar filas que contengan las llaves especificadas
    df_filtrado = df_filtrado[~df_filtrado['Descripción'].isin(llaves)]

    # Filtrar filas con valores nulos en la columna 'Cargos (PESOS)' para obtener los abonos
    df_abonos = df_filtrado[df_filtrado['Cargos (PESOS)'].isna()]

    # Filtrar filas con valores nulos en la columna 'Abonos (PESOS)' para obtener los cargos
    df_cargos = df_filtrado[df_filtrado['Abonos (PESOS)'].isna()]

    # Eliminar la columna 'Cargos (PESOS)' de los abonos
    df_abonos = df_abonos.drop("Cargos (PESOS)", axis=1)

    # Eliminar la columna 'Abonos (PESOS)' de los cargos
    df_cargos = df_cargos.drop("Abonos (PESOS)", axis=1)

    return df_abonos, df_cargos


def seleccionar_filas(df):
    filas_seleccionadas = []
    columnas = df.columns
    for index, row in df.iterrows():
        fecha = row[columnas[0]]
        monto = row[columnas[3]]
        descripcion = row[columnas[1]]
        # Borrar la consola antes de mostrar cada fila
        clear_console()
        print("¿Incluir?:\n")
        incluir = input(f"{fecha} ${monto:,} {descripcion}\n")

        if incluir == '':
            filas_seleccionadas.append(row)

    df_seleccionado = pd.DataFrame(filas_seleccionadas)

    return df_seleccionado

def main():
    # Leer el archivo de Excel en un DataFrame
    df = pd.read_excel('cartola.xls')

    # Leer el archivo "llaves.txt" y obtener las llaves
    with open('llaves.txt', 'r') as file:
        llaves = [line.strip() for line in file]

    # Filtrar el DataFrame y obtener los DataFrames de abonos y cargos
    df_abonos, df_cargos = filter_dataframe(df, llaves)

    # Crear una máscara booleana para las filas que contienen las llaves
    mask = df_cargos.iloc[:, 1].str.contains("|".join(llaves))

    # Filtrar el DataFrame para obtener las filas no removidas
    df_cargos = df_cargos[~mask]

    # Seleccionar filas para los abonos
    #df_abonos_seleccionados = seleccionar_filas(df_abonos)

    # Seleccionar filas para los cargos
    #df_cargos_seleccionados = seleccionar_filas(df_cargos)

    # Autorizar acceso a Google Sheets
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    gc = gspread.authorize(credentials)

    # Abrir la hoja de cálculo existente
    spreadsheet = gc.open("cuentas claras")

    # Establecer el idioma en español
    locale.setlocale(locale.LC_TIME, 'es_ES')

    # Obtener la fecha del primer dato del DataFrame
    fecha_primer_dato_str = df_cargos.iloc[0][0]

    # Convertir la fecha a formato datetime
    fecha_primer_dato = datetime.strptime(fecha_primer_dato_str, "%d/%m")

    # Crear el nombre de la hoja utilizando el formato deseado (por ejemplo, "Mes Año")
    hoja_titulo = fecha_primer_dato.strftime("%b 2023").capitalize() + " bot"

    # Verificar si la hoja ya existe
    try:
        worksheet = spreadsheet.worksheet(hoja_titulo)
    except gspread.exceptions.WorksheetNotFound:
        # La hoja no existe, crear una nueva
        worksheet = spreadsheet.add_worksheet(title=hoja_titulo, rows='100', cols='20')
    else:
        # La hoja ya existe, borrar su contenido
        worksheet.clear()

    df_cargos = df_cargos.iloc[:, [0, 1, 3]]
    data = [df_cargos.columns.values.tolist()] + df_cargos.values.tolist()

    # Insertar los datos en la hoja
    worksheet.update([df_cargos.columns.values.tolist()] + df_cargos.values.tolist())

    # Obtener la letra de la columna correspondiente a la columna 3
    letra_columna = chr(ord('A') + 2)  # La columna 3 es la columna 'C'

    ultima_fila = len(worksheet.col_values(1))
    # Crear la fórmula de suma para la columna 3
    formula_suma = f'=SUM({letra_columna}2:{letra_columna}{ultima_fila})'

    # Insertar una nueva fila al comienzo de la hoja
    worksheet.insert_row([''], 1)

    # Obtener la celda en la primera columna de la nueva fila
    celda_formula = worksheet.cell(1, 3)

    # Actualizar el valor de la celda con la fórmula de suma
    celda_formula.value = formula_suma

    # Guardar los cambios en la hoja de cálculo
    worksheet.update_cell(1, 3, celda_formula.value)

    print('Hoja creada o sobrescrita con éxito:', worksheet.title)

if __name__ == "__main__":
    main()
