import pandas as pd
import openpyxl
import re
import xlrd
import os
from forms.securitracForm import archivoSecuritrac
from forms.ubicomForm import archivoUbicom1, archivoUbicom2
from forms.MDVRForm import archivoMDVR1, archivoMDVR2
from forms.ubicarForm import archivoUbicar1, archivoUbicar2
from forms.wialonForm import archivoWialon1, archivoWialon2, archivoWialon3

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

def extraerIturan(file):
    # Cargar el archivo Excel usando xlrd
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)

    fecha_texto = sheet.cell_value(3, 0)  # La celda A4 es la cuarta fila (índice 3) y primera columna (índice 0)

    # Extraer la fecha del texto usando expresiones regulares
    fechas = re.findall(r'\d{2}/\d{2}/\d{4}', fecha_texto)
    fecha_inicio = fechas[0]  # Primera fecha encontrada

    # Lista para almacenar los datos extraídos
    datos_extraidos = []

    # Iterar sobre las filas del reporte
    for row_idx in range(7, sheet.nrows-1):
        placa = sheet.cell_value(row_idx, 2)
        km_recorridos = sheet.cell_value(row_idx, 8)
        num_excesos = int(sheet.cell_value(row_idx, 3))
        dia_trabajado = 1 if km_recorridos > 0 else 0
        preoperacional = 1 if dia_trabajado == 1 else None
        desplazamientos = sheet.cell_value(row_idx, 10)

        # Crear el diccionario con los datos de cada fila
        datos_fila = {
            'placa': placa,
            'fecha': fecha_inicio,
            'km_recorridos': km_recorridos,
            'dia_trabajado': dia_trabajado,
            'preoperacional': preoperacional,
            'num_excesos': num_excesos,
            'num_desplazamientos': desplazamientos,
        }

        # Añadir los datos de la fila a la lista
        datos_extraidos.append(datos_fila)

    return datos_extraidos

# Extraer los datos de los informes de MDVR.

def extraerMDVR(file1, file2): #file2 es el informe general, file2 es el informe de paradas (para determinar los desplazamientos)

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

def ejecutar_todas_extracciones(archivoMDVR1, archivoMDVR2, archivoIturan, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2):
    # Ejecutar cada función de extracción con los archivos proporcionados
    datosMDVR = extraerMDVR(archivoMDVR1, archivoMDVR2)
    datosIturan = extraerIturan(archivoIturan)
    datosSecuritrac = extraerSecuritrac(archivoSecuritrac)
    datosWialon = extraerWialon(archivoWialon1, archivoWialon2, archivoWialon3)
    datosUbicar = extraerUbicar(archivoUbicar1, archivoUbicar2)
    datosUbicom = extraerUbicom(archivoUbicom1, archivoUbicom2)

    # Unir todas las listas en una sola lista final
    listaFinal = datosMDVR + datosIturan + datosSecuritrac + datosWialon + datosUbicar + datosUbicom

    return listaFinal

# Crear el archivo Excel seguimiento.xlsx con los datos extraídos.

def crear_excel(archivoMDVR1, archivoMDVR2, archivoIturan, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2):
    data = ejecutar_todas_extracciones(archivoMDVR1, archivoMDVR2, archivoIturan, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2)
    df = pd.DataFrame(data)

    # Crear un DataFrame para el formato específico
    placas = df['placa'].unique()
    fechas = pd.date_range(start='2024-01-01', periods=366, freq='D')  # Rango de fechas para todo el año

    # Crear una estructura de DataFrame vacía con el formato deseado
    rows = []
    for placa in placas:
        rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Nº Excesos'})
        rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Nº Desplazamiento'})
        rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Día Trabajado'})
        rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Preoperacional'})
        rows.append({'PLACA': placa, 'SEGUIMIENTO': 'Km recorridos'})

    df_formato = pd.DataFrame(rows)

    # Agregar columnas para cada día del año
    for fecha in fechas:
        df_formato[fecha.strftime('%d/%m')] = ''

    # Rellenar los datos en el DataFrame con el formato deseado
    for _, row in df.iterrows():
        fecha = pd.to_datetime(row['fecha'], format='%d/%m/%Y')
        dia = fecha.strftime('%d/%m')
        df_formato.loc[(df_formato['PLACA'] == row['placa']) & (df_formato['SEGUIMIENTO'] == 'Nº Excesos'), dia] = row['num_excesos']
        df_formato.loc[(df_formato['PLACA'] == row['placa']) & (df_formato['SEGUIMIENTO'] == 'Nº Desplazamiento'), dia] = row['num_desplazamientos']
        df_formato.loc[(df_formato['PLACA'] == row['placa']) & (df_formato['SEGUIMIENTO'] == 'Día Trabajado'), dia] = row['dia_trabajado']
        df_formato.loc[(df_formato['PLACA'] == row['placa']) & (df_formato['SEGUIMIENTO'] == 'Preoperacional'), dia] = row['preoperacional']
        df_formato.loc[(df_formato['PLACA'] == row['placa']) & (df_formato['SEGUIMIENTO'] == 'Km recorridos'), dia] = row['km_recorridos']

    # Guardar el DataFrame en un archivo Excel
    writer = pd.ExcelWriter('seguimiento.xlsx', engine='openpyxl')
    df_formato.to_excel(writer, index=False, sheet_name='Sheet1')

    # Congelar paneles
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.freeze_panes = worksheet['C2']  # Congelar la primera fila y las dos primeras columnas

    writer.close()


# Definir ruta de los archivos de cada plataforma de manera universal.
archivoIturan = os.getcwd() + "\\outputIturan\\Over speed by vehicle (summary).xls"

# ESTO ES UN EJEMPLO DE COMO SE EJECUTA LA FUNCIÓN
crear_excel(
    archivoMDVR1,
    archivoMDVR2,
    archivoIturan,
    archivoSecuritrac,
    archivoWialon1,
    archivoWialon2,
    archivoWialon3,
    archivoUbicar1,
    archivoUbicar2,
    archivoUbicom1,
    archivoUbicom2
)



