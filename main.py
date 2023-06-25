import pandas as pd
import os

def clear_console():
    # Función para borrar la consola
    os.system('cls' if os.name == 'nt' else 'clear')

def filter_dataframe(df):
    # Eliminar las primeras 23 filas
    df = df.drop(df.index[:23])

    # Eliminar la primera y última columna
    df = df.iloc[:, 1:-1]

    # Asignar la primera fila como nombres de columna
    df.columns = df.iloc[0]

    # Eliminar la primera fila, ya que ahora son los nombres de columna
    df = df[1:]

    # Filtrar filas con valores no nulos en la columna 'Canal o Sucursal'
    df_filtrado = df.dropna(subset=['Canal o Sucursal'])

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

# Leer el archivo de Excel en un DataFrame
df = pd.read_excel('cartola.xls')

# Filtrar el DataFrame y obtener los DataFrames de abonos y cargos
df_abonos, df_cargos = filter_dataframe(df)

# Seleccionar filas para los abonos
df_abonos_seleccionados = seleccionar_filas(df_abonos)

# Seleccionar filas para los cargos
df_cargos_seleccionados = seleccionar_filas(df_cargos)

# Calcular la suma de los montos seleccionados en los abonos y cargos
total_abonos = df_abonos_seleccionados.iloc[:, 3].sum()
total_cargos = df_cargos_seleccionados.iloc[:, 3].sum()

# Imprimir los totales en pesos chilenos
print(f"Total Abonos ${total_abonos:,.2f} CLP")
print(f"Total Cargos ${total_cargos:,.2f} CLP")
