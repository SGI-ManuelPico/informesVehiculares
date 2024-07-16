import pandas as pd
import openpyxl
import re
import xlrd
import os
from openpyxl import load_workbook
from datetime import datetime 

# Extraer información del informe de Ubicar.

def extraerUbicar(file1, file2): # file 1 es el informe general, file2 es el informe de paradas (para determinar los desplazamientos)
    # Cargar el archivo de Excel
    workbook = openpyxl.load_workbook(file1)
    workbook2 = openpyxl.load_workbook(file2)

    sheet = workbook.active

    sheet2 = workbook2.active

    # Extraer la información necesaria del reporte
    placa_completa = sheet.cell(row=1, column=2).value
    # Extraer la parte relevante de la placa
    placa = placa_completa.split()[1] + placa_completa.split()[2]
    fecha = sheet.cell(row=2, column=2).value.split()[0].replace('-','/')
    km_recorridos = float(sheet.cell(row=4, column=2).value.replace(' Km', ''))
    dia_trabajado = 1 if km_recorridos > 0 else 0
    preoperacional = 1 if dia_trabajado == 1 else None
    numExcesos = sheet.cell(row=9, column=2).value

    #Contar número de desplazamientos
    num_desplazamientos = 0
    for i in range(5, sheet2.max_row + 1):
        if sheet2.cell(row=i, column=1).value == 'Movimiento':
            num_desplazamientos += 1




    # Crear el diccionario con los datos extraídos
    datos_extraidos = {
        'placa': placa,
        'fecha': fecha,
        'km_recorridos': km_recorridos,
        'dia_trabajado': dia_trabajado,
        'preoperacional': preoperacional,
        'num_excesos': numExcesos,
        'num_desplazamientos': num_desplazamientos,
    }

    return [datos_extraidos]

# Extraer los datos de los informes de Ituran.

def extraerIturan(file1, file2):
    # Cargar el archivo csv
    itu = pd.read_csv(file1)[['NICK_NAME', 'TOTAL_TRIP_DISTANCE', 'TOTAL_NUMBER_OF_TRIPS']]
    itu2 = pd.read_csv(file2)
    
    fecha = datetime.now().strftime('%d/%m/%Y')
    
    # Cambiar el nombre de las columnas
    itu = itu.rename(columns={
        'NICK_NAME': 'placa',
        'TOTAL_TRIP_DISTANCE': 'km_recorridos',
        'TOTAL_NUMBER_OF_TRIPS': 'num_desplazamientos'
    })
    
    # Agregar columna 'fecha'
    itu['fecha'] = fecha
    
    # Agregar columnas 'dia_trabajado' y 'preoperacional'
    itu['dia_trabajado'] = itu['km_recorridos'].apply(lambda x: 1 if x > 0 else 0)
    itu['preoperacional'] = itu['dia_trabajado'].apply(lambda x: 1 if x == 1 else '-')
    
    # Calcular el número de excesos de velocidad y crear DataFrame para fusiones
    excesos = itu2[itu2['TOP_SPEED'] > 80].groupby('V_NICK_NAME').size().reset_index(name='num_excesos')
    excesos = excesos.rename(columns={'V_NICK_NAME': 'placa'})
    
    # Unir el DataFrame de excesos con el DataFrame itu
    itu = itu.merge(excesos, on='placa', how='left')
    
    # Reemplazar los valores NaN en la columna num_excesos por 0
    itu['num_excesos'] = itu['num_excesos'].fillna(0).astype(int)
    
    # Convertir el DataFrame filtrado a un diccionario sin incluir el índice
    datos_extraidos = itu.to_dict(orient='records')
    
    return datos_extraidos

# Extraer los datos de los informes de MDVR.

def extraerMDVR(file1, file2): #file1 es el informe general, file2 es el informe de paradas (para determinar los desplazamientos)

# Cargar el archivo de Excel usando xlrd
    workbook = xlrd.open_workbook(file1)
    sheet = workbook.sheet_by_index(0)
    workbook2 = openpyxl.load_workbook(file2)
    sheet2 = workbook2.active


    # Extraer la información necesaria del reporte
    placa_completa = sheet.cell_value(0, 1)  # A1 es (0, 1)
    placa = placa_completa.replace('-', '')  # Quitar el guion de la placa
    fecha = sheet.cell_value(1, 1).split()[0]  # A2 es (1, 1)
    km_recorridos = float(sheet.cell_value(3, 1).replace(' Km', ''))  # A5 es (4, 1)
    dia_trabajado = 1 if km_recorridos > 0 else 0
    preoperacional = 1 if dia_trabajado == 1 else None
    num_excesos = int(sheet.cell_value(8, 1))  # A9 es (8, 1)

    #Contar número de desplazamientos
    num_desplazamientos = 0
    for i in range(1, sheet2.max_row + 1):
        if sheet2.cell(row=i, column=2).value == 'Movimiento':
            num_desplazamientos += 1

    # Crear el diccionario con los datos extraídos
    datos_extraidos = {
        'placa': placa,
        'fecha': fecha,
        'km_recorridos': km_recorridos,
        'dia_trabajado': dia_trabajado,
        'preoperacional': preoperacional,
        'num_excesos': num_excesos,
        'num_desplazamientos': num_desplazamientos,
    }

    return [datos_extraidos]

# Extraer los datos de los informes de Securitrac.

def extraerSecuritrac(file):
   # Cargar el archivo de Excel usando pandas
    df = pd.read_excel(file)

    # Diccionario para almacenar los datos por placa
    datos_por_placa = {}
    
    for index, row in df.iterrows():
        placa = row['NROMOVIL']
        evento = row['EVENTO']
        kilometros = float(row['KILOMETROS'])
        fecha = row['FECHAGPS']

    
        fecha_formateada = pd.to_datetime(fecha).strftime('%d/%m/%Y')

        if placa not in datos_por_placa:
            datos_por_placa[placa] = {
                'placa': placa,
                'fecha': fecha_formateada, 
                'km_recorridos': 0,
                'num_excesos': 0,
                'num_desplazamientos': 0
            }
        datos_por_placa[placa]['km_recorridos'] += kilometros
        if evento == 'Exc. Velocidad':
            datos_por_placa[placa]['num_excesos'] += 1
        datos_por_placa[placa]['num_desplazamientos'] += 1

    for placa, datos in datos_por_placa.items():
        datos['dia_trabajado'] = 1 if datos['km_recorridos'] > 0 else 0
        datos['preoperacional'] = 1 if datos['dia_trabajado'] == 1 else None

    datos = []
    for x in datos_por_placa.keys():
        datos.append(datos_por_placa[x])

    return datos

# Extraer los datos de los informes de Ubicom.

def extraerUbicom(file1, file2):
    # Cargar el archivo de Excel usando xlrd
    workbook = xlrd.open_workbook(file1)
    sheet = workbook.sheet_by_index(0)
    workbook2 = xlrd.open_workbook(file2)
    sheet2 = workbook2.sheet_by_index(0)

    # Extraer la información necesaria del reporte
    fecha = sheet.cell_value(11, 28).split()[0]  # Celda AC12
    
    km_recorridos = float(sheet.cell_value(20, 12))  # Celda M20
  
    num_excesos = int(sheet.cell_value(20, 21))  # Celda V20
  
    placa = sheet.cell_value(13, 24).split(' - ')[1].replace('(', '').replace(')', '') # Celda Y14, quitando el texto adicional
    

    # Calcular día trabajado y preoperacional
    dia_trabajado = 1 if km_recorridos > 0 else 0
    preoperacional = 1 if dia_trabajado == 1 else None

    # Extraer el número de desplazamientos
   
    num_desplazamientos = 0
    for i in range(17, sheet2.nrows - 1):  # Empezamos en la fila 18 y ajustamos para no contar la última fila
        if sheet2.cell_value(i, 10) != '':  # Columna K es el indice 10
            num_desplazamientos += 1

    # Crear el diccionario con los datos extraídos.
    datos_extraidos = {
        'placa': placa,
        'fecha': fecha,
        'km_recorridos': km_recorridos,
        'dia_trabajado': dia_trabajado,
        'preoperacional': preoperacional,
        'num_excesos': num_excesos,
        'num_desplazamientos': num_desplazamientos,
    }

    return [datos_extraidos]

# Extraer los datos de los informes de Wialon.

def extraerWialon(file1, file2, file3):
    datos_extraidos = []
    
    for file in [file1, file2, file3]:
        xl = pd.ExcelFile(file)
        
        # Extraer placa y fecha siempre
        if 'Statistics' in xl.sheet_names:
            statistics_df = xl.parse('Statistics', header=None)
            placa = statistics_df.iloc[0, 1]  # Celda B1
            fecha = statistics_df.iloc[1, 1].split()[0].replace('.', '/')  # Celda B2
            fecha_formateada = pd.to_datetime(fecha).strftime('%d/%m/%Y')
        
        
        # Verificar si el archivo tiene datos
        if 'Statistics' in xl.sheet_names and 'Excesos de velocidad' in xl.sheet_names and 'Cronología' in xl.sheet_names:
            km_recorridos = int(statistics_df.iloc[7, 1])  # Celda B8, quitando 'km'
            excesos_df = xl.parse('Excesos de velocidad', header=None)
            num_excesos = len(excesos_df) - 1  # Descontar la fila de encabezado
            crono = xl.parse('Cronología')
            
            # Extraer número de desplazamientos
            desplazamientos = 0
            for x in crono['Tipo'].to_list():
                if x == 'Trip':
                    desplazamientos += 1
            
            # Calcular día trabajado y preoperacional
            dia_trabajado = 1 if km_recorridos > 0 else 0
            preoperacional = 1 if dia_trabajado == 1 else None
            
            # Crear el diccionario con los datos extraídos
            datos = {
                'placa': placa,
                'fecha': fecha_formateada,
                'km_recorridos': km_recorridos,
                'dia_trabajado': dia_trabajado,
                'preoperacional': preoperacional,
                'num_excesos': num_excesos,
                'num_desplazamientos': desplazamientos
            }
        else:
            # Si el archivo no tiene datos, llenar con ceros pero usar placa y fecha
            datos = {
                'placa': placa,
                'fecha': fecha_formateada,
                'km_recorridos': 0,
                'dia_trabajado': 0,
                'preoperacional': None,
                'num_excesos': 0,
                'num_desplazamientos': 0
            }
        
        datos_extraidos.append(datos)
    
    return datos_extraidos

# Ejecuta todas las extracciones y las une en una única lista.

def ejecutar_todas_extracciones(archivoMDVR1, archivoMDVR2, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2):
    # Ejecutar cada función de extracción con los archivos proporcionados
    datosMDVR = extraerMDVR(archivoMDVR1, archivoMDVR2)
    datosIturan = extraerIturan(archivoIturan1 , archivoIturan2)
    datosSecuritrac = extraerSecuritrac(archivoSecuritrac)
    datosWialon = extraerWialon(archivoWialon1, archivoWialon2, archivoWialon3)
    datosUbicar = extraerUbicar(archivoUbicar1, archivoUbicar2)
    datosUbicom = extraerUbicom(archivoUbicom1, archivoUbicom2)

    # Unir todas las listas en una sola lista final
    listaFinal = datosMDVR + datosIturan + datosSecuritrac + datosWialon + datosUbicar + datosUbicom

    return listaFinal

# Crear el archivo Excel seguimiento.xlsx con los datos extraídos. Si el archivo ya existe, simplemente lo actualiza con los datos nuevos.

def crear_excel(mdvr_file1, mdvr_file2, ituran_file, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2, output_file='seguimiento.xlsx'):

    try:
        # Ejecutar todas las extracciones
        nuevos_datos = ejecutar_todas_extracciones(mdvr_file1, mdvr_file2, ituran_file, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2)
        
        # Convertir la lista de nuevos datos a DataFrame
        df_nuevos = pd.DataFrame(nuevos_datos)
        
        # Esto es para la persistencia
        if not os.path.exists(output_file):
            placas = df_nuevos['placa'].unique()
            fechas = pd.date_range(start='2024-01-01', periods=365, freq='D')  # Ajustar el rango de fechas según sea necesario
            
            rows = []
            for placa in placas:
                rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Nº Excesos'})
                rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Nº Desplazamiento'})
                rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Día Trabajado'})
                rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Preoperacional'})
                rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Km recorridos'})
            
            df_formato = pd.DataFrame(rows)
            
            # Agrega columnas para cada día del año con nombres de meses (No pude hacer que quedara tal cual como en el formato ejemplo)
            for fecha in fechas:
                mes_dia = fecha.strftime('%d/%m')
                df_formato[mes_dia] = ''
        else:
            print('Ya existe seguimiento.xlsx')
            # Leer el DataFrame existente
            df_formato = pd.read_excel(output_file)

            return df_formato

    finally:
    
    # Rellenar los datos en el DataFrame con el formato deseado
        for _, row in df_nuevos.iterrows():
            fecha = pd.to_datetime(row['fecha'], format='%d/%m/%Y')
            dia = fecha.strftime('%d/%m')
            placa = row['placa']
            
            df_formato.loc[(df_formato['PLACA'] == placa) & (df_formato['SEGUIMIENTO'] == 'Nº Excesos'), dia] = row['num_excesos']
            df_formato.loc[(df_formato['PLACA'] == placa) & (df_formato['SEGUIMIENTO'] == 'Nº Desplazamiento'), dia] = row['num_desplazamientos']
            df_formato.loc[(df_formato['PLACA'] == placa) & (df_formato['SEGUIMIENTO'] == 'Día Trabajado'), dia] = row['dia_trabajado']
            df_formato.loc[(df_formato['PLACA'] == placa) & (df_formato['SEGUIMIENTO'] == 'Preoperacional'), dia] = row['preoperacional']
            df_formato.loc[(df_formato['PLACA'] == placa) & (df_formato['SEGUIMIENTO'] == 'Km recorridos'), dia] = row['km_recorridos']
        
        # Guardar el DataFrame en un archivo Excel
        df_formato.to_excel(output_file, index=False)

# Infractores diario Ubicar


def infracUbicar(file1):
    # Leer el archivo de Excel y obtener las hojas
    df = pd.read_excel(file1, skiprows=2).iloc[:-1]  # Ignorar las primeras 4 filas y la última fila

    # Extraer la placa del vehículo de la celda B1
    placa = "JYT620"  # Reemplazar con la extracción correcta si se requiere

    # Convertir la columna 'Comienzo' a datetime y extraer solo la fecha y hora
    df['Comienzo'] = pd.to_datetime(df['Comienzo'], dayfirst= True)
    df['Fecha'] = df['Comienzo'].dt.strftime('%d/%m/%Y %H:%M:%S')
    
    # Convertir la columna 'Duración' a segundos
    def convertir_duracion_a_seg(duration):
        parts = duration.split(' ')
        total_seconds = 0
        for part in parts:
            if 'h' in part:
                total_seconds += int(part.replace('h', '')) * 3600
            elif 'min' in part:
                total_seconds += int(part.replace('min', '')) * 60
            elif 's' in part:
                total_seconds += int(part.replace('s', ''))
        return total_seconds

    df['Tiempo de Exceso'] = df['Duración'].apply(convertir_duracion_a_seg)
    
    # Crear el diccionario con el formato requerido
    registros = []
    for index, row in df.iterrows():
        registros.append({
            'PLACA': placa,
            'TIEMPO DE EXCESO': row['Tiempo de Exceso'],
            'DURACIÓN EN KM DE EXCESO': None,
            'VELOCIDAD MÁXIMA': float(row['Velocidad máxima'].replace(' kph', '')),
            'PROYECTO': None,
            'CONDUCTOR': None,
            'RUTA DE EXCESO': row['Posición'], # Por ahora guardamos las coordenadas porque pasarlas a dirección requiere de otras cosas.
            'FECHA': row['Fecha']
        })

    # Imprimir el número de registros
    print(f'Número de registros: {len(registros)}')
    
    # Retornar el diccionario de registros
    return registros

# Infractores diario MDVR

def infracMDVR(file):
    # Abrir el archivo Excel ignorando posibles corrupciones
    workbook = xlrd.open_workbook(file, ignore_workbook_corruption=True)
    
    # Leer el archivo Excel ignorando las primeras dos filas y la última fila
    df = pd.read_excel(workbook, skiprows=2).iloc[:-1]

    # Extraer la placa del vehículo de la celda B1
    placa = pd.read_excel(workbook).columns[1].replace(' ', '')

    # Convertir la columna 'Comienzo' a datetime y extraer la fecha y hora
    df['Comienzo'] = pd.to_datetime(df['Comienzo'], dayfirst=True)
    df['Fecha'] = df['Comienzo'].dt.strftime('%d/%m/%Y %H:%M:%S')

    # Dividir la columna 'Posición' en 'Latitud' y 'Longitud'
    df[['Latitud', 'Longitud']] = df['Posición'].str.split(',', expand=True)
    df['Latitud'] = df['Latitud'].astype(float)
    df['Longitud'] = df['Longitud'].astype(float)

    # Convertir el tiempo de exceso a segundos
    def convertirASegs(duration_str):
        parts = duration_str.split(' ')
        minutes = 0
        seconds = 0
        for part in parts:
            if 'min' in part:
                minutes += int(part.replace('min', ''))
            if 's' in part:
                seconds += int(part.replace('s', ''))
        return minutes * 60 + seconds

    df['Duración exceso de velocidad'] = df['Duración exceso de velocidad'].apply(convertirASegs)

    # Crear el diccionario en el formato requerido
    registros = []
    for index, row in df.iterrows():
        registro = {
            'PLACA': placa,
            'TIEMPO DE EXCESO': row['Duración exceso de velocidad'],
            'DURACIÓN EN KM DE EXCESO': None,
            'VELOCIDAD MÁXIMA': float(row['Velocidad máxima'].replace('kph', '')),
            'PROYECTO': '',
            'RUTA DE EXCESO': f"{row['Latitud']}, {row['Longitud']}",  # Por ahora guardamos las coordenadas porque pasarlas a dirección requiere de otras cosas.
            'CONDUCTOR': '',
            'FECHA': row['Fecha']
        }
        registros.append(registro)

    print(f"Número total de registros: {len(registros)}")
    return registros


# Infractores diario Securitrac

def infracSecuritrac(file):
    # Leer el archivo de Excel
    df = pd.read_excel(file)

    # Filtrar solo las filas con "Exc. Velocidad"
    df = df[df['EVENTO'] == 'Exc. Velocidad']

    # Crear el diccionario con el formato requerido
    registros = []
    for index, row in df.iterrows():
        fecha_formateada = pd.to_datetime(row['FECHAGPS']).strftime('%d/%m/%Y %H:%M:%S')
        registro = {
            'PLACA': row['NROMOVIL'],
            'TIEMPO DE EXCESO': None,
            'DURACIÓN EN KM DE EXCESO': None,
            'VELOCIDAD MÁXIMA': row['VELOCIDAD'],
            'RUTA DE EXCESO': row['POSICION'], 
            'PROYECTO': None,
            'CONDUCTOR': None,
            'FECHA': fecha_formateada,
        }
        registros.append(registro)

    print(f"Total de registros generados: {len(registros)}")
    print(registros[:5]) # Muestra los primeros 5 registros para verificar

    return registros

# Infractores diario Ituran

def infracIturan(file1):
    # Leer el archivo de Excel y obtener las hojas
    df = pd.read_csv(file1)
    
    # Filtrar la hoja que contiene la información relevante
    df_infracciones = df[['V_NICK_NAME', 'EVENT_DURATION_SEC', 'EVENT_DISTANCE', 'TOP_SPEED', 'VEHICLE_GROUP', 'ADDRESS', 'DRIVER_NAME', 'EVENT_START_DAY_TIME']]

    # Renombrar las columnas según lo solicitado
    df_infracciones.columns = ['PLACA', 'TIEMPO DE EXCESO', 'DURACIÓN EN KM DE EXCESO', 'VELOCIDAD MÁXIMA', 'PROYECTO', 'RUTA DE EXCESO', 'CONDUCTOR', 'FECHA']

    # Mantener la hora en la columna FECHA y convertirla al formato dd/mm/yyyy HH:MM:SS
    df_infracciones['FECHA'] = pd.to_datetime(df_infracciones['FECHA']).dt.strftime('%d/%m/%Y %H:%M:%S')

    # Filtrar los registros relevantes
    df_infracciones = df_infracciones[df_infracciones['DURACIÓN EN KM DE EXCESO'] > 0]

    num_registros = len(df_infracciones)
    print(f"Número de registros: {num_registros}")

    # Convertir el DataFrame en un diccionario
    datos_infracciones = df_infracciones.to_dict(orient='records')

    return datos_infracciones

# Infractores diario Wialon

def infracWialon(file1):
    xls = pd.ExcelFile(file1)

    # Verificar si la hoja 'Excesos de velocidad' existe
    if 'Excesos de velocidad' not in xls.sheet_names:
        print("La hoja 'Excesos de velocidad' no está presente en el archivo.")
        return []

    # Leer la hoja 'Statistics' para obtener la placa
    df_stats = pd.read_excel(xls, 'Statistics')
    placa = df_stats.columns[1]  # Celda B1 en la hoja "Statistics"

    

    # Leer la hoja 'Excesos de velocidad' para obtener la información necesaria
    df_excesos = pd.read_excel(xls, 'Excesos de velocidad')

    # Convertir la columna 'Comienzo' a datetime y extraer solo la fecha y hora
    df_excesos['Comienzo'] = pd.to_datetime(df_excesos['Comienzo'])
    df_excesos['Fecha'] = df_excesos['Comienzo'].dt.strftime('%d/%m/%Y %H:%M:%S')

    # Convertir la columna 'Duración' a segundos
    def convertirDuracionASegundos(duration_str):
        parts = duration_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        else:
            return int(parts[0])

    df_excesos['Duración en segundos'] = df_excesos['Duración'].apply(convertirDuracionASegundos)

    # Crear el diccionario con el formato requerido
    registros = []
    for index, row in df_excesos.iterrows():
        registro = {
            'PLACA': placa,
            'TIEMPO DE EXCESO': row['Duración en segundos'],
            'DURACIÓN EN KM DE EXCESO': '',  # No disponible en este documento
            'VELOCIDAD MÁXIMA': row['Velocidad máxima'],
            'PROYECTO': '',  # No disponible en este documento
            'CONDUCTOR': '',  # No disponible en este documento
            'RUTA DE EXCESO': row['Localización'],
            'FECHA': row['Fecha']
        }
        registros.append(registro)

    print(f"Total de registros extraídos: {len(registros)}")
    return registros# Extraer información infractores Wialon

# Ejecuta todas las extracciones y las une en una única lista.

def infracTodos(fileIturan, fileMDVR, fileUbicar, fileWialon, fileWialon2, fileWialon3, fileSecuritrac):
    # Ejecutar cada función infrac y obtener los resultados
    registros_ituran = infracIturan(fileIturan)
    registros_mdvr = infracMDVR(fileMDVR)
    registros_ubicar = infracUbicar(fileUbicar)
    registros_wialon = infracWialon(fileWialon)
    registros_wialon2 = infracWialon(fileWialon2)
    registros_wialon3 = infracWialon(fileWialon3)
    registros_securitrac = infracSecuritrac(fileSecuritrac)

    # Combinar todos los resultados en una sola lista
    todos_registros = (
        registros_ituran + registros_mdvr + registros_ubicar + registros_wialon + registros_wialon2 + registros_wialon3 +registros_securitrac
    )

    print(f"Total de registros combinados: {len(todos_registros)}")
    return todos_registros

# Actualiza la hoja Infractores de la hoja de Excel. (Todavía falta testear esta función con otros archivos)

def actualizarInfractores(file_seguimiento, file_Ituran, file_MDVR, file_Ubicar, file_Wialon, file_Securitrac):
    # Obtener todas las infracciones combinadas
    todos_registros = infracTodos(file_Ituran, file_MDVR, file_Ubicar, file_Wialon, file_Securitrac)
    df_infractores = pd.DataFrame(todos_registros)

    # Convertir la columna 'FECHA' a datetime y luego a string con el formato correcto
    df_infractores['FECHA'] = pd.to_datetime(df_infractores['FECHA'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y %H:%M:%S')

    # Cargar el archivo existente y añadir una nueva hoja
    with pd.ExcelWriter(file_seguimiento, engine='openpyxl', mode='a') as writer:
        try:
            # Cargar el libro de trabajo existente
            writer.book = load_workbook(file_seguimiento)
            # Verificar si la hoja 'Infractores' ya existe
            if 'Infractores' in writer.book.sheetnames:
                # Leer la hoja existente en un DataFrame
                df_existente = pd.read_excel(file_seguimiento, sheet_name='Infractores')
                # Concatenar los datos existentes con los nuevos datos
                df_final = pd.concat([df_existente, df_infractores], ignore_index=True)
            else:
                # Si la hoja no existe, simplemente usar los nuevos datos
                df_final = df_infractores

            # Escribir el DataFrame en la hoja 'Infractores'
            df_final.to_excel(writer, sheet_name='Infractores', index=False)
        except Exception as e:
            print(f"Error al actualizar el archivo Excel: {e}")

# Odómetro Ituran

def OdomIturan(file):
    # Leer el archivo de Excel
    od = pd.read_excel(file)

    # Extraer la placa y el odómetro

    df = od[['V_PLATE_NUMBER', 'END_ODOMETER']]

    # Renombrar las columnas

    df.columns = ['PLACA', 'KILOMETRAJE']

    # Crear el diccionario con el formato requerido

    datos = df.to_dict('records')
    return datos

# Odómetro Ubicar 

def odomUbicar(file):
    # Leer el archivo de Excel
    df = pd.read_excel(file)  # Ajustar skiprows si es necesario

    # Extraer la placa del vehículo de la celda B1, si agregan otro carro a esta plataforma toca cambiar como se extrae esto.
    placa = 'JYT620'

    # Extraer el odómetro de la celda correspondiente
    odometro = df.iloc[11, 2]  # Ajustar el índice si es necesario

    # Crear el diccionario con el formato requerido
    registro = {
        'PLACA': placa,
        'KILOMETRAJE': float(odometro.split()[0].replace(',', ''))
    }

    return [registro]

# Actualizar hoja de Odómetro en el Excel

def actualizarOdom(file_seguimiento, file_ituran, file_ubicar):
    # Obtener todos los odómetros combinados
    todos_registros = OdomIturan(file_ituran) + odomUbicar(file_ubicar)
    df_odometros = pd.DataFrame(todos_registros)

    # Cargar el archivo existente y añadir una nueva hoja, sobreescribiendo si ya existe
    with pd.ExcelWriter(file_seguimiento, engine='openpyxl', mode='a') as writer:
        # Verificar si la hoja ya existe
        if 'Odómetro' in writer.book.sheetnames:
            # Eliminar la hoja existente
            std = writer.book['Odometro']
            writer.book.remove(std)
        # Escribir el DataFrame en una nueva hoja llamada 'Odometro'
        df_odometros.to_excel(writer, sheet_name='Odómetro', index=False)

mdvr_file1 = r"C:\Users\SGI SAS\Downloads\general_information_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720804327.xls"
mdvr_file2 = r"C:\Users\SGI SAS\Downloads\stops_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720804333.xlsx"
archivoIturan1 = r"C:\Users\SGI SAS\Downloads\report.csv"
archivoIturan2 = r"C:\Users\SGI SAS\Downloads\report(1).csv"
securitrac_file = r"C:\Users\SGI SAS\Downloads\exported-excel.xls"
wialon_file1 = r"C:\Users\SGI SAS\Downloads\LPN816_INFORME_GENERAL_TM_V1.0_2024-07-12_16-30-46.xlsx"
wialon_file2 = r"C:\Users\SGI SAS\Downloads\LPN821_INFORME_GENERAL_TM_V1.0_2024-07-12_16-30-57.xlsx"
wialon_file3 = r"C:\Users\SGI SAS\Downloads\JTV645_INFORME_GENERAL_TM_V1.0_2024-07-12_16-30-28.xlsx"
ubicar_file1 = r"C:\Users\SGI SAS\Downloads\general_information_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720803683.xlsx"
ubicar_file2 = r"C:\Users\SGI SAS\Downloads\stops_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720803694.xlsx"
ubicom_file1 = r"C:\Users\SGI SAS\Downloads\ReporteDiario.xls"
ubicom_file2 = r"C:\Users\SGI SAS\Downloads\Estacionados.xls"



crear_excel(mdvr_file1, mdvr_file2, archivoIturan1, archivoIturan2, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2, output_file='seguimiento.xlsx')

## ACTUALIZAR HOJAS DE INDICADORES 

