import pandas as pd

# Cargar el archivo Excel de entrada
archivo_entrada = 'segmento_1.xlsx'
df = pd.read_excel(archivo_entrada)

# Procesar los datos
datos_seleccionados = df[['Nombre']]

# Crear un nuevo archivo Excel con los datos seleccionados
archivo_salida = 'hola.xlsx'
datos_seleccionados.to_excel(archivo_salida, index=False)

print(f'Datos procesados y guardados en {archivo_salida}')
