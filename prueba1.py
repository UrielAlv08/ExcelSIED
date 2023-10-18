import pandas as pd

# Cargar el archivo Excel de entrada
archivo_entrada = 'prueba1.xlsx'
df = pd.read_excel(archivo_entrada)

# Eliminar las columnas en blanco
df = df[['Nombre', 'Contrato', 'TotalCredito', 'Fecha1', 'Plazos', 'Pendiente', 'Pagado', 'ValorTasa', 'Periodo', 'FechaInicioP', 'SaldoIn', 'Amortizacion', 'Ordinarios', 'IVA', 'PagoTot', 'PeriodoPagado', 'FechaPago']]

# Cambiar los nombres de columna de origen
nombres_origen = ['Nombre', 'Contrato', 'TotalCredito', 'Fecha1', 'Plazos', 'Pendiente', 'Pagado', 'ValorTasa', 'Periodo', 'FechaInicioP', 'SaldoIn', 'Amortizacion', 'Ordinarios', 'IVA', 'PagoTot', 'PeriodoPagado', 'FechaPago']
df.columns = nombres_origen

# Crear un nuevo archivo Excel con los datos seleccionados sin índices
archivo_salida = 'estructura_nueva4.xlsx'
df.to_excel(archivo_salida, index=False)

print(f'Datos procesados y guardados en {archivo_salida}')
