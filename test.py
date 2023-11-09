import pandas as pd

# Lee el archivo Excel
ruta_archivo_excel = "nombre_del_archivo.xlsx"  # Reemplaza con el nombre de tu archivo Excel
data = pd.read_excel(ruta_archivo_excel)

# Convierte los valores de la columna de tiempo a objetos de tiempo
data['Tiempo'] = pd.to_datetime(data['Tiempo'])

# Crea una nueva columna que combine la hora y el minuto
data['Hora_y_Minuto'] = data['Tiempo'].dt.strftime('%H:%M')

# Crea un DataFrame vacío para almacenar los resultados finales
resultados = pd.DataFrame(columns=['Distribuidor', 'Branch', 'Horario Mas Repetido', 'Envios Variando Mas de 3 Minutos', 'Cantidad Total de Envios'])

# Realiza el análisis desglosado por cada distribuidor y sucursal
for (distribuidor, sucursal), grupo in data.groupby(['Distribuidor', 'Branch']):
    # Encuentra el horario más repetido para el distribuidor y sucursal actual
    horario_mas_repetido = grupo['Hora_y_Minuto'].mode().values[0]

    # Convierte el horario más repetido nuevamente a formato datetime para obtener la hora y el minuto
    hora_mas_repetida, minuto_mas_repetido = map(int, horario_mas_repetido.split(':'))

    # Filtra los envíos que varían más de 3 minutos con el envío más repetido
    envios_varian_mas_3_minutos = grupo[abs(grupo['Tiempo'].dt.hour - hora_mas_repetida) > 0]  # Compara la hora
    envios_varian_mas_3_minutos = envios_varian_mas_3_minutos[abs(envios_varian_mas_3_minutos['Tiempo'].dt.minute - minuto_mas_repetido) > 3]  # Compara el minuto

    # Cuenta el número de envíos que varían más de 3 minutos para el distribuidor y sucursal actual
    cantidad_envios_varian_mas_3_minutos = len(envios_varian_mas_3_minutos)

    # Obtiene la cantidad total de envíos para el distribuidor y sucursal actual
    cantidad_total_envios = len(grupo)

    # Agrega los resultados al DataFrame de resultados
    resultados = resultados.append({
        'Distribuidor': distribuidor,
        'Branch': sucursal,
        'Horario Mas Repetido': horario_mas_repetido,
        'Envios Variando Mas de 3 Minutos': cantidad_envios_varian_mas_3_minutos,
        'Cantidad Total de Envios': cantidad_total_envios
    }, ignore_index=True)

    # Agrega las columnas adicionales al archivo original
    data.loc[(data['Distribuidor'] == distribuidor) & (data['Branch'] == sucursal), 'Horario Mas Repetido'] = horario_mas_repetido
    data.loc[(data['Distribuidor'] == distribuidor) & (data['Branch'] == sucursal), 'Envio Varia Mas de 3 Minutos'] = abs(data['Tiempo'].dt.hour - hora_mas_repetida) * 60 + abs(data['Tiempo'].dt.minute - minuto_mas_repetido) > 3

# Guarda los resultados en un nuevo archivo Excel
ruta_resultados_excel = "resultados.xlsx"  # Reemplaza con el nombre que deseas para el archivo Excel de resultados
resultados.to_excel(ruta_resultados_excel, index=False)

# Guarda el archivo original con las columnas adicionales
ruta_archivo_original_con_columnas = "archivo_original_con_columnas.xlsx"  # Reemplaza con el nombre que deseas para el archivo Excel con columnas adicionales
data.to_excel(ruta_archivo_original_con_columnas, index=False)

print("Análisis finalizado. Resultados guardados en el archivo:", ruta_resultados_excel)
print("Archivo original con columnas adicionales guardado en el archivo:", ruta_archivo_original_con_columnas)
