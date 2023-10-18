import pandas as pd

# Cargar el archivo Excel de entrada
archivo_entrada = 'segmento_1.xlsx'
df = pd.read_excel(archivo_entrada)

# Procesar los datos (supongamos que deseas extraer las columnas 'Nombre' y 'Edad')
columnas_seleccionadas = ['Nombre']  # Agrega todas las columnas que necesitas

# Eliminar filas vacías
df = df.dropna(subset=columnas_seleccionadas, how='all')

datos_seleccionados = df[columnas_seleccionadas]

# Crear un nuevo archivo Excel con los datos seleccionados sin índices
archivo_salida = 'holi.xlsx'
datos_seleccionados.to_excel(archivo_salida, index=False)

print(f'Datos procesados y guardados en {archivo_salida}')
