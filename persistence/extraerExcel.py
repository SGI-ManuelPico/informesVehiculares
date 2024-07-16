import pandas as pd
import openpyxl
import re
import xlrd
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from datetime import datetime
from openpyxl.styles import Font, PatternFill

class ExtraerExcel():
    """
        Clase para realizar todas las funciones respectivas a la extracción de datos de un excel
    """

    def __init__(self):
        pass

    def ituran(self, accion: int, file1, file2):
        """
        Si accion == 1: Es para la función que ejecuta la extracción para la base de datos. SQL
        Si accion == 0: Es para la función que ejecuta la extracción para el excel. EXCEL
        """
        if accion == 1:
            # Cargar el archivo csv
            itu = pd.read_csv(file1)[['NICK_NAME', 'TOTAL_TRIP_DISTANCE', 'TOTAL_NUMBER_OF_TRIPS']]
            itu2 = pd.read_csv(file2)
            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")

            fecha = datetime.now().strftime('%d/%m/%Y') + ' ' + current_time
            
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
            itu['preoperacional'] = itu['dia_trabajado'].apply(lambda x: 1 if x == 1 else 0)
            
            # Calcular el número de excesos de velocidad y crear DataFrame para fusiones
            excesos = itu2[itu2['TOP_SPEED'] > 80].groupby('V_NICK_NAME').size().reset_index(name='num_excesos')
            excesos = excesos.rename(columns={'V_NICK_NAME': 'placa'})
            
            # Unir el DataFrame de excesos con el DataFrame itu
            itu = itu.merge(excesos, on='placa', how='left')
            
            # Reemplazar los valores NaN en la columna num_excesos por 0
            itu['num_excesos'] = itu['num_excesos'].fillna(0).astype(int)

            itu['proveedor'] = 'Ituran'
            
            # Convertir el DataFrame filtrado a un diccionario sin incluir el índice
            datos_extraidos = itu.to_dict(orient='records')
            
            return datos_extraidos
        elif accion == 0:
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
            itu['preoperacional'] = itu['dia_trabajado'].apply(lambda x: 1 if x == 1 else 0)
            
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

    def MDVR(self, accion: int, file1, file2): #file2 es el informe general, file2 es el informe de paradas (para determinar los desplazamientos)
        """
        Si accion == 1: Es para la función que ejecuta la extracción para la base de datos. SQL
        Si accion == 0: Es para la función que ejecuta la extracción para el excel. EXCEL
        """
        if accion == 1:
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
            preoperacional = 1 if dia_trabajado == 1 else 0
            num_excesos = int(sheet.cell_value(8, 1))  # A9 es (8, 1)

            #Contar número de desplazamientos
            num_desplazamientos = 0
            for i in range(1, sheet2.max_row + 1):
                if sheet2.cell(row=i, column=2).value == 'Movimiento':
                    num_desplazamientos += 1
            
            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")

            # Crear el diccionario con los datos extraídos
            datos_extraidos = {
                'placa': placa,
                'fecha': fecha + ' ' + current_time,
                'km_recorridos': km_recorridos,
                'dia_trabajado': dia_trabajado,
                'preoperacional': preoperacional,
                'num_excesos': num_excesos,
                'num_desplazamientos': num_desplazamientos,
                'proveedor': 'MDVR'
            }
            return [datos_extraidos]
        elif accion == 0:
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

    def ubicar(self, accion: int, file1, file2): # file 1 es el informe general, file2 es el informe de paradas (para determinar los desplazamientos)
        """
        Si accion == 1: Es para la función que ejecuta la extracción para la base de datos. SQL
        Si accion == 0: Es para la función que ejecuta la extracción para el excel. EXCEL
        """
        if accion == 1:
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
            preoperacional = 1 if dia_trabajado == 1 else 0
            numExcesos = sheet.cell(row=9, column=2).value

            #Contar número de desplazamientos
            num_desplazamientos = 0
            for i in range(5, sheet2.max_row + 1):
                if sheet2.cell(row=i, column=1).value == 'Movimiento':
                    num_desplazamientos += 1

            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")

            # Crear el diccionario con los datos extraídos
            datos_extraidos = {
                'placa': placa,
                'fecha': fecha + '' + current_time,
                'km_recorridos': km_recorridos,
                'dia_trabajado': dia_trabajado,
                'preoperacional': preoperacional,
                'num_excesos': numExcesos,
                'num_desplazamientos': num_desplazamientos,
                'proveedor': 'Ubicar'
            }

            return [datos_extraidos]
        elif accion == 0:
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

    def ubicom(self, accion: int, file1, file2):
        """
        Si accion == 1: Es para la función que ejecuta la extracción para la base de datos. SQL
        Si accion == 0: Es para la función que ejecuta la extracción para el excel. EXCEL
        """
        if accion == 1:
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
            preoperacional = 1 if dia_trabajado == 1 else 0

            # Extraer el número de desplazamientos
        
            num_desplazamientos = 0
            for i in range(17, sheet2.nrows - 1):  # Empezamos en la fila 18 y ajustamos para no contar la última fila
                if sheet2.cell_value(i, 10) != '':  # Columna K es el indice 10
                    num_desplazamientos += 1

            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")

            # Crear el diccionario con los datos extraídos
            datos_extraidos = {
                'placa': placa,
                'fecha': fecha + ' ' + current_time,
                'km_recorridos': km_recorridos,
                'dia_trabajado': dia_trabajado,
                'preoperacional': preoperacional,
                'num_excesos': num_excesos,
                'num_desplazamientos': num_desplazamientos,
                'proveedor': 'Ubicom'
            }

            return [datos_extraidos]
        elif accion == 0:
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

    def securitrac(self, accion: int, file):
        """
        Si accion == 1: Es para la función que ejecuta la extracción para la base de datos. SQL
        Si accion == 0: Es para la función que ejecuta la extracción para el excel. EXCEL
        """
        if accion == 1:
            # Cargar el archivo de Excel usando pandas
            df = pd.read_excel(file)

            # Diccionario para almacenar los datos por placa
            datos_por_placa = {}

            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")
            
            for index, row in df.iterrows():
                placa = row['NROMOVIL']
                evento = row['EVENTO']
                kilometros = float(row['KILOMETROS'])
                fecha = row['FECHAGPS']

            
                fecha_formateada = pd.to_datetime(fecha).strftime('%d/%m/%Y')

                if placa not in datos_por_placa:
                    datos_por_placa[placa] = {
                        'placa': placa,
                        'fecha': fecha_formateada + ' ' + current_time, 
                        'km_recorridos': 0,
                        'num_excesos': 0,
                        'num_desplazamientos': 0,
                        'proveedor': 'Securitrac'
                    }
                datos_por_placa[placa]['km_recorridos'] += kilometros
                if evento == 'Exc. Velocidad':
                    datos_por_placa[placa]['num_excesos'] += 1
                datos_por_placa[placa]['num_desplazamientos'] += 1

            for placa, datos in datos_por_placa.items():
                datos['dia_trabajado'] = 1 if datos['km_recorridos'] > 0 else 0
                datos['preoperacional'] = 1 if datos['dia_trabajado'] == 1 else 0

            datos = []
            for x in datos_por_placa.keys():
                datos.append(datos_por_placa[x])

            return datos
        elif accion == 0:
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

    def wialon(self, accion: int, file1, file2, file3):
        """
        Si accion == 1: Es para la función que ejecuta la extracción para la base de datos. SQL
        Si accion == 0: Es para la función que ejecuta la extracción para el excel. EXCEL
        """
        if accion == 1:
            datos_extraidos = []

            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")
            
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
                    preoperacional = 1 if dia_trabajado == 1 else 0
                    
                    # Crear el diccionario con los datos extraídos
                    datos = {
                        'placa': placa,
                        'fecha': fecha_formateada + ' ' + current_time,
                        'km_recorridos': km_recorridos,
                        'dia_trabajado': dia_trabajado,
                        'preoperacional': preoperacional,
                        'num_excesos': num_excesos,
                        'num_desplazamientos': desplazamientos,
                        'proveedor': 'Wialon'
                    }

                else:
                    # Si el archivo no tiene datos, llenar con ceros pero usar placa y fecha
                    datos = {
                        'placa': placa,
                        'fecha': fecha_formateada + ' ' + current_time,
                        'km_recorridos': 0,
                        'dia_trabajado': 0,
                        'preoperacional': 0,
                        'num_excesos': 0,
                        'num_desplazamientos': 0,
                        'proveedor': 'Wialon'
                    }
                
                datos_extraidos.append(datos)
            
            return datos_extraidos
        elif accion == 0:
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

    def ejecutarTodasExtraccionesSQL(self, file_ituran1, file_ituran2, file_MDVR1, file_MDVR2, file_Ubicar1, file_Ubicar2, file_Ubicom1, file_Ubicom2, file_Securitrac, file_Wialon1, file_Wialon2, file_Wialon3):
        """
        Funcion de extraccion para la base de datos. SQL
        """
        # Ejecutar cada función de extracción con los archivos proporcionados
        datos_mdvr = self.MDVR(1, file_MDVR1, file_MDVR2)
        datos_ituran = self.ituran(1, file_ituran1, file_ituran2)
        datos_securitrac = self.securitrac(1, file_Securitrac)
        datos_wialon = self.wialon(1, file_Wialon1, file_Wialon2, file_Wialon3)
        datos_ubicar = self.ubicar(1, file_Ubicar1, file_Ubicar2)
        datos_ubicom = self.ubicom(1, file_Ubicom1, file_Ubicom2)

        # Unir todas las listas en una sola lista final
        lista_final = datos_ituran + datos_mdvr + datos_ubicar + datos_ubicom + datos_securitrac + datos_wialon 

        df_final = pd.DataFrame(lista_final)

        df_final.rename(columns={
        'placa': 'placa',
        'fecha': 'fecha',
        'km_recorridos': 'kmRecorridos',
        'dia_trabajado': 'diaTrabajado',
        'preoperacional': 'preoperacional',
        'num_excesos': 'numExcesos',
        'num_desplazamientos': 'numDesplazamientos',
        'proveedor': 'proveedor'
        }, inplace=True)

        ordenColumnas = [
        'placa',
        'kmRecorridos',
        'numDesplazamientos',
        'diaTrabajado',
        'preoperacional',
        'numExcesos',
        'proveedor',
        'fecha'
        ]

        # Reordenar las columnas de df_final
        df_final = df_final[ordenColumnas]

        return df_final
    
    def ejecutarTodasExtraccionesExcel(self, archivoMDVR1, archivoMDVR2, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2):
        """
        Funcion de extraccion para el excel. EXCEL
        """
        # Ejecuta todas las extracciones y las une en una única lista.
        # Ejecutar cada función de extracción con los archivos proporcionados
        datosMDVR = self.MDVR(0, archivoMDVR1, archivoMDVR2)
        datosIturan = self.ituran(0, archivoIturan1, archivoIturan2)
        datosSecuritrac = self.securitrac(0, archivoSecuritrac)
        datosWialon = self.wialon(0, archivoWialon1, archivoWialon2, archivoWialon3)
        datosUbicar = self.ubicar(0, archivoUbicar1, archivoUbicar2)
        datosUbicom = self.ubicom(0, archivoUbicom1, archivoUbicom2)

        # Unir todas las listas en una sola lista final
        listaFinal = datosMDVR + datosIturan + datosSecuritrac + datosWialon + datosUbicar + datosUbicom

        return listaFinal
    
    # Infractores diario Ubicar
    def infracUbicar(self, file1):
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
    def infracMDVR(self, file):
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
    def infracSecuritrac(self, file):
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
    def infracIturan(self, file1):
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
    def infracWialon(self, file1):
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
    def infracTodos(self, fileIturan, fileMDVR, fileUbicar, fileWialon, fileWialon2, fileWialon3, fileSecuritrac):
        # Ejecutar cada función infrac y obtener los resultados
        registros_ituran = self.infracIturan(fileIturan)
        registros_mdvr = self.infracMDVR(fileMDVR)
        registros_ubicar = self.infracUbicar(fileUbicar)
        registros_wialon = self.infracWialon(fileWialon)
        registros_wialon2 = self.infracWialon(fileWialon2)
        registros_wialon3 = self.infracWialon(fileWialon3)
        registros_securitrac = self.infracSecuritrac(fileSecuritrac)

        # Combinar todos los resultados en una sola lista
        todos_registros = (
            registros_ituran + registros_mdvr + registros_ubicar + registros_wialon + registros_wialon2 + registros_wialon3 +registros_securitrac
        )

        print(f"Total de registros combinados: {len(todos_registros)}")
        return todos_registros
    
    # Odómetro Ituran
    def OdomIturan(file):
        # Leer el archivo de Excel
        od = pd.read_csv(file)

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
    
    # Estas funciones solo son pequeñas modificaciones de las de extracción de los datos históricos para llenar la tabla de seguimiento con los datos históricos
    def ituranHistoricoSQL(self, file1, file2):
        # Leer los archivos Excel
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Convertir las fechas a un formato de fecha datetime
        df1['DT_DRIVE_DATE'] = pd.to_datetime(df1['DT_DRIVE_DATE']).dt.date
        df2['EVENT_START_DAY_TIME'] = pd.to_datetime(df2['EVENT_START_DAY_TIME']).dt.date

        # Crear un DataFrame para almacenar los resultados
        results = []

        # Obtener las placas únicas
        placas = df1['V_NICK_NAME'].unique()

        for placa in placas:
            # Filtrar los datos por placa
            df1_placa = df1[df1['V_NICK_NAME'] == placa]
            df2_placa = df2[df2['V_NICK_NAME'] == placa]

            # Agrupar por fecha para contar los desplazamientos y sumar las distancias
            daily_stats_df1 = df1_placa.groupby('DT_DRIVE_DATE').agg({
                'TRIP_DISTANCE': 'sum',
                'DT_DRIVE_DATE': 'count'  # Esta cuenta el número de desplazamientos
            }).rename(columns={'DT_DRIVE_DATE': 'num_desplazamientos', 'TRIP_DISTANCE': 'km_recorridos'}).reset_index()

            # Agrupar por fecha para contar los excesos
            daily_stats_df2 = df2_placa.groupby('EVENT_START_DAY_TIME').size().reset_index(name='num_excesos')

            # Unir ambos DataFrames por fecha
            daily_stats = pd.merge(daily_stats_df1, daily_stats_df2, left_on='DT_DRIVE_DATE', right_on='EVENT_START_DAY_TIME', how='left').fillna(0)

            current_datetime = datetime.now()
            current_time = current_datetime.strftime("%H:%M:%S")

            for _, row in daily_stats.iterrows():
                fecha_formateada = row['DT_DRIVE_DATE'].strftime('%d/%m/%Y')
                datos_extraidos = {
                    'placa': placa,
                    'fecha': fecha_formateada + ' ' + current_time,
                    'km_recorridos': round(row['km_recorridos'], 2),
                    'num_desplazamientos': row['num_desplazamientos'],
                    'dia_trabajado': 1 if row['km_recorridos'] > 0 else 0,
                    'preoperacional': 1 if row['km_recorridos'] > 0 else None,
                    'num_excesos': int(row['num_excesos']),
                    'proveedor': 'Ituran'
                }
                results.append(datos_extraidos)

        return results

    def MDVRHistoricoSQL(self, file1, file2):
        # Leer el archivo de desplazamientos
        df1 = pd.read_excel(file1, engine='xlrd', header=2, skipfooter=2)
        df1['Fecha'] = pd.to_datetime(df1['Fecha'], format='%d/%m/%Y %H:%M:%S').dt.date
        
        # Limpiar y convertir 'Longitud de ruta' a float
        df1['Longitud de ruta'] = df1['Longitud de ruta'].str.replace(' Km', '').astype(float)
        
        # Contar desplazamientos y sumar kilómetros por día
        desplazamientos_km = df1.groupby('Fecha').agg(
            num_desplazamientos=('Fecha', 'size'),
            km_recorridos=('Longitud de ruta', 'sum')
        ).reset_index()
        
        
        file_path = file2
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        # Extraer los datos de la hoja y convertirlos en un DataFrame
        data = sheet.values
        next(data) 
        next(data) 

        next(data) 
        

        cols = ["Comienzo", "Fin", "Duración exceso de velocidad", "Velocidad máxima", "Velocidad media", "Posición"]
        data = list(data)
        df2 = pd.DataFrame(data, columns=cols)
        
        # Quitar la última fila
        df2 = df2.iloc[:-1]
        
        # Convertir la columna 'Comienzo' a datetime
        df2['Comienzo'] = pd.to_datetime(df2['Comienzo'], format='%d/%m/%Y %H:%M:%S').dt.date
        
        # Contar excesos de velocidad por día
        excesos_velocidad = df2.groupby('Comienzo').size().reset_index(name='num_excesos')
        
        # Unir los datos de desplazamientos y excesos de velocidad
        datos_historicos = pd.merge(desplazamientos_km, excesos_velocidad, left_on='Fecha', right_on='Comienzo', how='left').fillna(0)
        
        # Añadir la columna fija 'placa'
        datos_historicos['placa'] = 'KSZ298'
        
        # Añadir las columnas faltantes
        datos_historicos['dia_trabajado'] = (datos_historicos['km_recorridos'] > 0).astype(int)
        datos_historicos['preoperacional'] = datos_historicos['dia_trabajado'].replace(0, '-')
        
        # Seleccionar y renombrar columnas
        datos_historicos = datos_historicos[['placa', 'Fecha', 'km_recorridos', 'num_desplazamientos', 'dia_trabajado', 'preoperacional', 'num_excesos']]
        datos_historicos = datos_historicos.rename(columns={'Fecha': 'fecha'})
        
        # Convertir la columna 'fecha' a datetime
        datos_historicos['fecha'] = pd.to_datetime(datos_historicos['fecha'])

        current_datetime = datetime.now()
        current_time = current_datetime.strftime("%H:%M:%S")
        
        # Convertir la fecha a formato dd/mm/yyyy
        datos_historicos['fecha'] = datos_historicos['fecha'].dt.strftime('%d/%m/%Y') + ' ' + current_time

        # Asignar una nueva columna 'proveedor' con todas las entradas como 'Wialon'
        datos_historicos['proveedor'] = 'MDVR'
        
        return datos_historicos.to_dict(orient='records')

    def ubicarHistoricoSQL(self, file1, file2):
        # Leer el archivo de desplazamientos
        df1 = pd.read_excel(file1, engine='xlrd', header=2, skipfooter=2)
        df1['Fecha'] = df1['Fecha'].apply(lambda x: x.replace('-', '/'))
        df1['Fecha'] = pd.to_datetime(df1['Fecha'], format="%d/%m/%Y %H:%M:%S").dt.date

        # Limpiar y convertir 'Longitud de ruta' a float
        df1['Longitud de ruta'] = df1['Longitud de ruta'].str.replace(' Km', '').astype(float)

        # Contar desplazamientos y sumar kilómetros por día
        desplazamientos_km = df1.groupby('Fecha').agg(
            num_desplazamientos=('Fecha', 'size'),
            km_recorridos=('Longitud de ruta', 'sum')
        ).reset_index()

        # Leer el archivo de excesos de velocidad
        workbook = openpyxl.load_workbook(file2)
        sheet = workbook.active

        # Extraer los datos de la hoja y convertirlos en un DataFrame
        data = sheet.values
        next(data) # Saltar la primera fila de datos (encabezados principales)
        next(data) # Saltar la segunda fila de datos (encabezados secundarios)
        next(data)

        cols = ["Comienzo", "Fin", "Duración exceso de velocidad", "Velocidad máxima", "Velocidad media", "Posición"]
        data = list(data)
        df2 = pd.DataFrame(data, columns=cols)

        # Quitar la última fila
        df2 = df2.iloc[:-1]

        # Convertir la columna 'Comienzo' a datetime
        df2['Comienzo'] = pd.to_datetime(df2['Comienzo'].str.replace('.', '/'), format="%d/%m/%Y %H:%M:%S", errors='coerce').dt.date

        # Contar excesos de velocidad por día
        excesos_velocidad = df2.groupby('Comienzo').size().reset_index(name='num_excesos')

        # Unir los datos de desplazamientos y excesos de velocidad
        datos_historicos = pd.merge(desplazamientos_km, excesos_velocidad, left_on='Fecha', right_on='Comienzo', how='left').fillna(0)

        # Añadir la columna fija 'placa'
        datos_historicos['placa'] = 'JYT682'

        # Añadir las columnas faltantes
        datos_historicos['dia_trabajado'] = (datos_historicos['km_recorridos'] > 0).astype(int)
        datos_historicos['preoperacional'] = (datos_historicos['km_recorridos'] > 0).astype(int)
        datos_historicos.loc[datos_historicos['dia_trabajado'] == 0, 'preoperacional'] = '-'

        # Seleccionar y renombrar columnas
        datos_historicos = datos_historicos[['placa', 'Fecha', 'km_recorridos', 'num_desplazamientos', 'dia_trabajado', 'preoperacional', 'num_excesos']]
        datos_historicos = datos_historicos.rename(columns={'Fecha': 'fecha'})

        # Convertir la columna 'fecha' a datetime
        datos_historicos['fecha'] = pd.to_datetime(datos_historicos['fecha'])

        current_datetime = datetime.now()
        current_time = current_datetime.strftime("%H:%M:%S")

        # Convertir la fecha a formato dd/mm/yyyy
        datos_historicos['fecha'] = datos_historicos['fecha'].dt.strftime('%d/%m/%Y') + ' ' + current_time

        # Asignar una nueva columna 'proveedor' con todas las entradas como 'Wialon'
        datos_historicos['proveedor'] = 'Ubicar'

        return datos_historicos.to_dict(orient='records')

    def ubicomHistoricoSQL(self, file1, file2):
        # Leer el archivo de reporte diario
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Limpiar y convertir las columnas relevantes
        df1['Fecha'] = pd.to_datetime(df1['Fecha'], format='%Y-%m-%d').dt.date
        df1['Distancia recorrida'] = df1['Distancia recorrida'].astype(float)
        df1['Número de excesos de velocidad'] = df1['Número de excesos de velocidad'].astype(int)

        df2['Fecha'] = pd.to_datetime(df2['Fecha'], format='%Y-%m-%d').dt.date
        df2['Número'] = df2['Número'].astype(int)

        # Agrupar por fecha para sumar los kilómetros recorridos, contar los excesos de velocidad y los desplazamientos
        reporte_diario = df1.groupby('Fecha').agg({
            'Distancia recorrida': 'sum',
            'Número de excesos de velocidad': 'sum'
        }).rename(columns={
            'Distancia recorrida': 'km_recorridos',
            'Número de excesos de velocidad': 'num_excesos'
        }).reset_index()

        desplazamientos_diarios = df2.groupby('Fecha').agg({
            'Número': 'sum'
        }).rename(columns={'Número': 'num_desplazamientos'}).reset_index()

        # Unir los dataframes
        reporte_completo = pd.merge(reporte_diario, desplazamientos_diarios, on='Fecha', how='left')

        # Añadir la columna fija 'placa'
        reporte_completo['placa'] = 'FNM236'

        # Añadir columnas 'preoperacional' y 'día trabajado'
        reporte_completo['preoperacional'] = 1
        reporte_completo['dia_trabajado'] = 1

        # Si los km recorridos son igual a 0, entonces el número de desplazamientos también debe serlo y 'día trabajado' debe ser 0
        reporte_completo.loc[reporte_completo['km_recorridos'] == 0, 'num_desplazamientos'] = 0
        reporte_completo.loc[reporte_completo['km_recorridos'] == 0, 'dia_trabajado'] = 0
        reporte_completo.loc[reporte_completo['km_recorridos'] == 0, 'preoperacional'] = '-'

        # Seleccionar y renombrar columnas
        reporte_completo = reporte_completo[['placa', 'Fecha', 'km_recorridos', 'num_excesos', 'num_desplazamientos', 'preoperacional', 'dia_trabajado']]
        reporte_completo = reporte_completo.rename(columns={'Fecha': 'fecha'})

        # Convertir la columna 'fecha' a datetime
        reporte_completo['fecha'] = pd.to_datetime(reporte_completo['fecha'])

        current_datetime = datetime.now()
        current_time = current_datetime.strftime("%H:%M:%S")

        # Convertir la fecha a formato dd/mm/yyyy
        reporte_completo['fecha'] = reporte_completo['fecha'].dt.strftime('%d/%m/%Y') + ' ' + current_time

        # Asignar una nueva columna 'proveedor' con todas las entradas como 'Wialon'
        reporte_completo['proveedor'] = 'Ubicom'

        return reporte_completo.to_dict(orient='records')

    def securitracHistoricoSQL(self, file):
        # Leer el archivo de Excel
        df = pd.read_excel(file)

        # Convertir la columna FECHAGPS a datetime
        df['FECHAGPS'] = pd.to_datetime(df['FECHAGPS'])

        # Extraer la fecha sin la hora
        df['Fecha'] = df['FECHAGPS'].dt.date

        # Calcular los kilómetros recorridos por día y placa, tomando el valor máximo
        km_recorridos = df.groupby(['NROMOVIL', 'Fecha'])['KILOMETROS'].max().reset_index()

        # Contar los excesos de velocidad por día y placa
        excesos_velocidad = df[df['EVENTO'] == 'Exc. Velocidad'].groupby(['NROMOVIL', 'Fecha']).size().reset_index(name='num_excesos')

        # Contar los desplazamientos por día y placa
        num_desplazamientos = df.groupby(['NROMOVIL', 'Fecha']).size().reset_index(name='num_desplazamientos')

        # Unir los datos de kilómetros recorridos, excesos de velocidad y desplazamientos
        datos_historicos = km_recorridos.merge(excesos_velocidad, on=['NROMOVIL', 'Fecha'], how='left').fillna(0)
        datos_historicos = datos_historicos.merge(num_desplazamientos, on=['NROMOVIL', 'Fecha'], how='left').fillna(0)

        # Añadir la columna fija 'placa'
        datos_historicos['placa'] = datos_historicos['NROMOVIL']

        # Añadir la columna día trabajado (1 si km_recorridos > 0, si no 0)
        datos_historicos['dia_trabajado'] = datos_historicos['KILOMETROS'].apply(lambda x: 1 if x > 0 else 0)

        # Añadir la columna preoperacional (1 si dia_trabajado es 1, None de lo contrario)
        datos_historicos['preoperacional'] = datos_historicos['dia_trabajado'].apply(lambda x: 1 if x == 1 else '-')

        # Seleccionar y renombrar columnas necesarias
        datos_historicos = datos_historicos[['placa', 'Fecha', 'KILOMETROS', 'num_excesos', 'num_desplazamientos', 'preoperacional', 'dia_trabajado']]
        datos_historicos = datos_historicos.rename(columns={'Fecha': 'fecha', 'KILOMETROS': 'km_recorridos'})
        
        current_datetime = datetime.now()
        current_time = current_datetime.strftime("%H:%M:%S")

        # Convertir la columna fecha a formato dd/mm/yyyy
        datos_historicos['fecha'] = pd.to_datetime(datos_historicos['fecha']).dt.strftime('%d/%m/%Y') + ' ' + current_time

        # Asignar una nueva columna 'proveedor' con todas las entradas como 'Wialon'
        datos_historicos['proveedor'] = 'Securitrac'

        return datos_historicos.to_dict(orient='records')

    def wialonHistoricoSQL(self, file_path):
        # Leer el archivo de Excel y obtener las hojas
        xls = pd.ExcelFile(file_path)
        
        # Leer la hoja 'Statistics' para obtener la placa
        df_stats = pd.read_excel(xls, 'Statistics')
        placa = df_stats.columns[1]  # Celda B1 en la hoja "Statistics"
        
        # Leer la hoja 'Excesos de velocidad' para contar los excesos por día
        df_excesos = pd.read_excel(xls, 'Excesos de velocidad')
        df_excesos['Comienzo'] = pd.to_datetime(df_excesos['Comienzo'])
        df_excesos['fecha'] = df_excesos['Comienzo'].dt.date
        num_excesos = df_excesos.groupby('fecha').size().reset_index(name='num_excesos')
        
        # Leer la hoja 'Cronología' para contar los desplazamientos por día
        df_cronologia = pd.read_excel(xls, 'Cronología')
        df_cronologia['Comienzo'] = pd.to_datetime(df_cronologia['Comienzo'])
        df_cronologia['fecha'] = df_cronologia['Comienzo'].dt.date
        num_desplazamientos = df_cronologia[df_cronologia['Tipo'] == 'Trip'].groupby('fecha').size().reset_index(name='num_desplazamientos')
        
        # Leer la hoja 'Calles visitadas' para sumar el kilometraje por día
        df_calles = pd.read_excel(xls, 'Calles visitadas')
        df_calles['Comienzo'] = pd.to_datetime(df_calles['Comienzo'])
        df_calles['fecha'] = df_calles['Comienzo'].dt.date
        df_calles['Kilometraje'] = round(df_calles['Kilometraje'].astype(str).str.replace(' km', '').astype(float), 2)
        km_recorridos = df_calles.groupby('fecha')['Kilometraje'].sum().reset_index(name='km_recorridos')
        
        # Combinar los datos de excesos, desplazamientos y kilometraje
        datos_historicos = pd.merge(num_excesos, num_desplazamientos, on='fecha', how='outer').fillna(0)
        datos_historicos = pd.merge(datos_historicos, km_recorridos, on='fecha', how='outer').fillna(0)
        
        # Añadir las columnas adicionales
        datos_historicos['placa'] = placa
        datos_historicos['dia_trabajado'] = (datos_historicos['num_desplazamientos'] > 0).astype(int)
        datos_historicos['preoperacional'] = datos_historicos['dia_trabajado'].apply(lambda x: 1 if x == 1 else '-')
        
        # Seleccionar y renombrar columnas
        datos_historicos = datos_historicos[['placa', 'fecha', 'km_recorridos', 'num_excesos', 'num_desplazamientos', 'preoperacional', 'dia_trabajado']]
        datos_historicos.rename(columns={'fecha': 'Fecha'}, inplace=True)

        current_datetime = datetime.now()
        current_time = current_datetime.strftime("%H:%M:%S")

        
        # Convertir la fecha a formato dd/mm/yyyy
        datos_historicos['Fecha'] = pd.to_datetime(datos_historicos['Fecha']).dt.strftime('%d/%m/%Y') + ' ' + current_time


        # Asignar una nueva columna 'proveedor' con todas las entradas como 'Wialon'
        datos_historicos['proveedor'] = 'Wialon'
        
        return datos_historicos.to_dict(orient='records')
    
    # Ituran
    def histIturanExcel(self, file1, file2):
        # Leer los archivos Excel
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Convertir las fechas a un formato de fecha datetime
        df1['DT_DRIVE_DATE'] = pd.to_datetime(df1['DT_DRIVE_DATE']).dt.date
        df2['EVENT_START_DAY_TIME'] = pd.to_datetime(df2['EVENT_START_DAY_TIME']).dt.date

        # Crear un DataFrame para almacenar los resultados
        results = []

        # Obtener las placas únicas
        placas = df1['V_NICK_NAME'].unique()

        for placa in placas:
            # Filtrar los datos por placa
            df1_placa = df1[df1['V_NICK_NAME'] == placa]
            df2_placa = df2[df2['V_NICK_NAME'] == placa]

            # Agrupar por fecha para contar los desplazamientos y sumar las distancias
            daily_stats_df1 = df1_placa.groupby('DT_DRIVE_DATE').agg({
                'TRIP_DISTANCE': 'sum',
                'DT_DRIVE_DATE': 'count'  # Esta cuenta el número de desplazamientos
            }).rename(columns={'DT_DRIVE_DATE': 'num_desplazamientos', 'TRIP_DISTANCE': 'km_recorridos'}).reset_index()

            # Agrupar por fecha para contar los excesos
            daily_stats_df2 = df2_placa.groupby('EVENT_START_DAY_TIME').size().reset_index(name='num_excesos')

            # Unir ambos DataFrames por fecha
            daily_stats = pd.merge(daily_stats_df1, daily_stats_df2, left_on='DT_DRIVE_DATE', right_on='EVENT_START_DAY_TIME', how='left').fillna(0)

            for _, row in daily_stats.iterrows():
                fecha_formateada = row['DT_DRIVE_DATE'].strftime('%d/%m/%Y')
                datos_extraidos = {
                    'placa': placa,
                    'fecha': fecha_formateada,
                    'km_recorridos': round(row['km_recorridos'], 2),
                    'num_desplazamientos': row['num_desplazamientos'],
                    'dia_trabajado': 1 if row['km_recorridos'] > 0 else 0,
                    'preoperacional': 1 if row['km_recorridos'] > 0 else None,
                    'num_excesos': int(row['num_excesos'])
                }
                results.append(datos_extraidos)

        return results

    # MDVR
    def histMDVRExcel(self, file1, file2):
        # Leer el archivo de desplazamientos
        df1 = pd.read_excel(file1, engine='xlrd', header=2, skipfooter=2)
        df1['Fecha'] = pd.to_datetime(df1['Fecha'], format='%d/%m/%Y %H:%M:%S').dt.date
        
        # Limpiar y convertir 'Longitud de ruta' a float
        df1['Longitud de ruta'] = df1['Longitud de ruta'].str.replace(' Km', '').astype(float)
        
        # Contar desplazamientos y sumar kilómetros por día
        desplazamientos_km = df1.groupby('Fecha').agg(
            num_desplazamientos=('Fecha', 'size'),
            km_recorridos=('Longitud de ruta', 'sum')
        ).reset_index()
        
        
        file_path = file2
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        # Extraer los datos de la hoja y convertirlos en un DataFrame
        data = sheet.values
        next(data) 
        next(data) 

        next(data) 
        

        cols = ["Comienzo", "Fin", "Duración exceso de velocidad", "Velocidad máxima", "Velocidad media", "Posición"]
        data = list(data)
        df2 = pd.DataFrame(data, columns=cols)
        
        # Quitar la última fila
        df2 = df2.iloc[:-1]
        
        # Convertir la columna 'Comienzo' a datetime
        df2['Comienzo'] = pd.to_datetime(df2['Comienzo'], format='%d/%m/%Y %H:%M:%S').dt.date
        
        # Contar excesos de velocidad por día
        excesos_velocidad = df2.groupby('Comienzo').size().reset_index(name='num_excesos')
        
        # Unir los datos de desplazamientos y excesos de velocidad
        datos_historicos = pd.merge(desplazamientos_km, excesos_velocidad, left_on='Fecha', right_on='Comienzo', how='left').fillna(0)
        
        # Añadir la columna fija 'placa'
        datos_historicos['placa'] = 'KSZ298'
        
        # Añadir las columnas faltantes
        datos_historicos['dia_trabajado'] = (datos_historicos['km_recorridos'] > 0).astype(int)
        datos_historicos['preoperacional'] = datos_historicos['dia_trabajado'].replace(0, None)
        
        # Seleccionar y renombrar columnas
        datos_historicos = datos_historicos[['placa', 'Fecha', 'km_recorridos', 'num_desplazamientos', 'dia_trabajado', 'preoperacional', 'num_excesos']]
        datos_historicos = datos_historicos.rename(columns={'Fecha': 'fecha'})
        
        # Convertir la columna 'fecha' a datetime
        datos_historicos['fecha'] = pd.to_datetime(datos_historicos['fecha'])
        
        # Convertir la fecha a formato dd/mm/yyyy
        datos_historicos['fecha'] = datos_historicos['fecha'].dt.strftime('%d/%m/%Y')
        
        return datos_historicos.to_dict(orient='records')

    # Ubicar
    def histUbicarExcel(self, file1, file2):
        # Leer el archivo de desplazamientos
        df1 = pd.read_excel(file1, engine='xlrd', header=2, skipfooter=2)
        df1['Fecha'] = df1['Fecha'].apply(lambda x: x.replace('-', '/'))
        df1['Fecha'] = pd.to_datetime(df1['Fecha'], format="%d/%m/%Y %H:%M:%S").dt.date

        # Limpiar y convertir 'Longitud de ruta' a float
        df1['Longitud de ruta'] = df1['Longitud de ruta'].str.replace(' Km', '').astype(float)

        # Contar desplazamientos y sumar kilómetros por día
        desplazamientos_km = df1.groupby('Fecha').agg(
            num_desplazamientos=('Fecha', 'size'),
            km_recorridos=('Longitud de ruta', 'sum')
        ).reset_index()

        # Leer el archivo de excesos de velocidad
        workbook = openpyxl.load_workbook(file2)
        sheet = workbook.active

        # Extraer los datos de la hoja y convertirlos en un DataFrame
        data = sheet.values
        next(data) # Saltar la primera fila de datos (encabezados principales)
        next(data) # Saltar la segunda fila de datos (encabezados secundarios)
        next(data)

        cols = ["Comienzo", "Fin", "Duración exceso de velocidad", "Velocidad máxima", "Velocidad media", "Posición"]
        data = list(data)
        df2 = pd.DataFrame(data, columns=cols)

        # Quitar la última fila
        df2 = df2.iloc[:-1]

        # Convertir la columna 'Comienzo' a datetime
        df2['Comienzo'] = pd.to_datetime(df2['Comienzo'].str.replace('.', '/'), format="%d/%m/%Y %H:%M:%S", errors='coerce').dt.date

        # Contar excesos de velocidad por día
        excesos_velocidad = df2.groupby('Comienzo').size().reset_index(name='num_excesos')

        # Unir los datos de desplazamientos y excesos de velocidad
        datos_historicos = pd.merge(desplazamientos_km, excesos_velocidad, left_on='Fecha', right_on='Comienzo', how='left').fillna(0)

        # Añadir la columna fija 'placa'
        datos_historicos['placa'] = 'JYT682'

        # Añadir las columnas faltantes
        datos_historicos['dia_trabajado'] = (datos_historicos['km_recorridos'] > 0).astype(int)
        datos_historicos['preoperacional'] = (datos_historicos['km_recorridos'] > 0).astype(int)
        datos_historicos.loc[datos_historicos['dia_trabajado'] == 0, 'preoperacional'] = None

        # Seleccionar y renombrar columnas
        datos_historicos = datos_historicos[['placa', 'Fecha', 'km_recorridos', 'num_desplazamientos', 'dia_trabajado', 'preoperacional', 'num_excesos']]
        datos_historicos = datos_historicos.rename(columns={'Fecha': 'fecha'})

        # Convertir la columna 'fecha' a datetime
        datos_historicos['fecha'] = pd.to_datetime(datos_historicos['fecha'])

        # Convertir la fecha a formato dd/mm/yyyy
        datos_historicos['fecha'] = datos_historicos['fecha'].dt.strftime('%d/%m/%Y')

        return datos_historicos.to_dict(orient='records')

    # Ubicom 
    def histUbicomExcel(self, file1, file2):
        # Leer el archivo de reporte diario
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Limpiar y convertir las columnas relevantes
        df1['Fecha'] = pd.to_datetime(df1['Fecha'], format='%Y-%m-%d').dt.date
        df1['Distancia recorrida'] = df1['Distancia recorrida'].astype(float)
        df1['Número de excesos de velocidad'] = df1['Número de excesos de velocidad'].astype(int)

        df2['Fecha'] = pd.to_datetime(df2['Fecha'], format='%Y-%m-%d').dt.date
        df2['Número'] = df2['Número'].astype(int)

        # Agrupar por fecha para sumar los kilómetros recorridos, contar los excesos de velocidad y los desplazamientos
        reporte_diario = df1.groupby('Fecha').agg({
            'Distancia recorrida': 'sum',
            'Número de excesos de velocidad': 'sum'
        }).rename(columns={
            'Distancia recorrida': 'km_recorridos',
            'Número de excesos de velocidad': 'num_excesos'
        }).reset_index()

        desplazamientos_diarios = df2.groupby('Fecha').agg({
            'Número': 'sum'
        }).rename(columns={'Número': 'num_desplazamientos'}).reset_index()

        # Unir los dataframes
        reporte_completo = pd.merge(reporte_diario, desplazamientos_diarios, on='Fecha', how='left')

        # Añadir la columna fija 'placa'
        reporte_completo['placa'] = 'FNM236'

        # Añadir columnas 'preoperacional' y 'día trabajado'
        reporte_completo['preoperacional'] = 1
        reporte_completo['dia_trabajado'] = 1

        # Si los km recorridos son igual a 0, entonces el número de desplazamientos también debe serlo y 'día trabajado' debe ser 0
        reporte_completo.loc[reporte_completo['km_recorridos'] == 0, 'num_desplazamientos'] = 0
        reporte_completo.loc[reporte_completo['km_recorridos'] == 0, 'dia_trabajado'] = 0
        reporte_completo.loc[reporte_completo['km_recorridos'] == 0, 'preoperacional'] = None

        # Seleccionar y renombrar columnas
        reporte_completo = reporte_completo[['placa', 'Fecha', 'km_recorridos', 'num_excesos', 'num_desplazamientos', 'preoperacional', 'dia_trabajado']]
        reporte_completo = reporte_completo.rename(columns={'Fecha': 'fecha'})

        # Convertir la columna 'fecha' a datetime
        reporte_completo['fecha'] = pd.to_datetime(reporte_completo['fecha'])

        # Convertir la fecha a formato dd/mm/yyyy
        reporte_completo['fecha'] = reporte_completo['fecha'].dt.strftime('%d/%m/%Y')

        return reporte_completo.to_dict(orient='records')

    # Securitrac
    def histSecuritracExcel(self, file):

        # Leer el archivo de Excel
        df = pd.read_excel(file)

        # Convertir la columna FECHAGPS a datetime
        df['FECHAGPS'] = pd.to_datetime(df['FECHAGPS'])

        # Extraer la fecha sin la hora
        df['Fecha'] = df['FECHAGPS'].dt.date

        # Calcular los kilómetros recorridos por día y placa, tomando el valor máximo
        km_recorridos = df.groupby(['NROMOVIL', 'Fecha'])['KILOMETROS'].max().reset_index()

        # Contar los excesos de velocidad por día y placa
        excesos_velocidad = df[df['EVENTO'] == 'Exc. Velocidad'].groupby(['NROMOVIL', 'Fecha']).size().reset_index(name='num_excesos')

        # Contar los desplazamientos por día y placa
        num_desplazamientos = df.groupby(['NROMOVIL', 'Fecha']).size().reset_index(name='num_desplazamientos')

        # Unir los datos de kilómetros recorridos, excesos de velocidad y desplazamientos
        datos_historicos = km_recorridos.merge(excesos_velocidad, on=['NROMOVIL', 'Fecha'], how='left').fillna(0)
        datos_historicos = datos_historicos.merge(num_desplazamientos, on=['NROMOVIL', 'Fecha'], how='left').fillna(0)

        # Añadir la columna fija 'placa'
        datos_historicos['placa'] = datos_historicos['NROMOVIL']

        # Añadir la columna día trabajado (1 si km_recorridos > 0, si no 0)
        datos_historicos['dia_trabajado'] = datos_historicos['KILOMETROS'].apply(lambda x: 1 if x > 0 else 0)

        # Añadir la columna preoperacional (1 si dia_trabajado es 1, None de lo contrario)
        datos_historicos['preoperacional'] = datos_historicos['dia_trabajado'].apply(lambda x: 1 if x == 1 else None)

        # Seleccionar y renombrar columnas necesarias
        datos_historicos = datos_historicos[['placa', 'Fecha', 'KILOMETROS', 'num_excesos', 'num_desplazamientos', 'preoperacional', 'dia_trabajado']]
        datos_historicos = datos_historicos.rename(columns={'Fecha': 'fecha', 'KILOMETROS': 'km_recorridos'})

        # Convertir la columna fecha a formato dd/mm/yyyy
        datos_historicos['fecha'] = pd.to_datetime(datos_historicos['fecha']).dt.strftime('%d/%m/%Y')

        return datos_historicos.to_dict(orient='records')

    # Wialon
    def histWialonExcel(self, file_path):
        # Leer el archivo de Excel y obtener las hojas
        xls = pd.ExcelFile(file_path)
        
        # Leer la hoja 'Statistics' para obtener la placa
        df_stats = pd.read_excel(xls, 'Statistics')
        placa = df_stats.columns[1]  # Celda B1 en la hoja "Statistics"
        
        # Leer la hoja 'Excesos de velocidad' para contar los excesos por día
        df_excesos = pd.read_excel(xls, 'Excesos de velocidad')
        df_excesos['Comienzo'] = pd.to_datetime(df_excesos['Comienzo'])
        df_excesos['fecha'] = df_excesos['Comienzo'].dt.date
        num_excesos = df_excesos.groupby('fecha').size().reset_index(name='num_excesos')
        
        # Leer la hoja 'Cronología' para contar los desplazamientos por día
        df_cronologia = pd.read_excel(xls, 'Cronología')
        df_cronologia['Comienzo'] = pd.to_datetime(df_cronologia['Comienzo'])
        df_cronologia['fecha'] = df_cronologia['Comienzo'].dt.date
        num_desplazamientos = df_cronologia[df_cronologia['Tipo'] == 'Trip'].groupby('fecha').size().reset_index(name='num_desplazamientos')
        
        # Leer la hoja 'Calles visitadas' para sumar el kilometraje por día
        df_calles = pd.read_excel(xls, 'Calles visitadas')
        df_calles['Comienzo'] = pd.to_datetime(df_calles['Comienzo'])
        df_calles['fecha'] = df_calles['Comienzo'].dt.date
        df_calles['Kilometraje'] = round(df_calles['Kilometraje'].astype(str).str.replace(' km', '').astype(float), 2)
        km_recorridos = df_calles.groupby('fecha')['Kilometraje'].sum().reset_index(name='km_recorridos')
        
        # Combinar los datos de excesos, desplazamientos y kilometraje
        datos_historicos = pd.merge(num_excesos, num_desplazamientos, on='fecha', how='outer').fillna(0)
        datos_historicos = pd.merge(datos_historicos, km_recorridos, on='fecha', how='outer').fillna(0)
        
        # Añadir las columnas adicionales
        datos_historicos['placa'] = placa
        datos_historicos['dia_trabajado'] = (datos_historicos['num_desplazamientos'] > 0).astype(int)
        datos_historicos['preoperacional'] = datos_historicos['dia_trabajado'].apply(lambda x: 1 if x == 1 else None)
        
        # Seleccionar y renombrar columnas
        datos_historicos = datos_historicos[['placa', 'fecha', 'km_recorridos', 'num_excesos', 'num_desplazamientos', 'preoperacional', 'dia_trabajado']]
        datos_historicos.rename(columns={'fecha': 'Fecha'}, inplace=True)
        
        # Convertir la fecha a formato dd/mm/yyyy
        datos_historicos['Fecha'] = pd.to_datetime(datos_historicos['Fecha']).dt.strftime('%d/%m/%Y')
        
        return datos_historicos.to_dict(orient='records')
    
    # Para correr esta parte de infractores del script necesitamos file_ituran2, fileMDVR2, file_ubicar2, file_Wailon (1, 2, 3), file_securitrac. 
    # Nota: Todos estos archivos corresponden a los de excesos de velocidad. 
    def infracIturanHistorico(self, file1):
        # Leer el archivo de Excel y obtener las hojas
        df = pd.read_excel(file1)
        
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

    #MDVR
    # Función para convertir a segundos,
    def convert_to_seconds(self, duration_str):
        
        parts = duration_str.split()
        minutes = 0
        seconds = 0
        for part in parts:
            if 'min' in part:
                minutes = int(part.replace('min', ''))
            elif 's' in part:
                seconds = int(part.replace('s', ''))
        return minutes * 60 + seconds

    def infracMDVRHistorico(self, file1):
        # Leer el archivo Excel ignorando las primeras dos filas y la última fila
        df = pd.read_excel(file1, skiprows=2).iloc[:-1]
        
        # Extraer la placa del vehículo de la celda B1
        placa = pd.read_excel(file1).columns[1].replace('-', '')
        
        # Convertir la columna 'Comienzo' a datetime y extraer solo la fecha y hora
        df['Comienzo'] = pd.to_datetime(df['Comienzo'], dayfirst= True)
        df['Fecha'] = df['Comienzo'].dt.strftime('%d/%m/%Y %H:%M:%S')
        
        # Dividir la columna 'Posición' en 'Latitud' y 'Longitud'
        df[['Latitud', 'Longitud']] = df['Posición'].str.split(',', expand=True)
        df['Latitud'] = df['Latitud'].astype(float)
        df['Longitud'] = df['Longitud'].astype(float)
        
        # Convertir el tiempo de exceso a segundos
        df['Duración exceso de velocidad'] = df['Duración exceso de velocidad'].apply(self.convert_to_seconds)
        
        # Crear el diccionario con el formato requerido
        registros = []
        for index, row in df.iterrows():
            registro = {
                'PLACA': placa,
                'TIEMPO DE EXCESO': row['Duración exceso de velocidad'],
                'DURACIÓN EN KM DE EXCESO': '',
                'VELOCIDAD MÁXIMA': float(row['Velocidad máxima'].replace('kph', '')),
                'PROYECTO': '',
                'RUTA DE EXCESO': f"{row['Latitud']}, {row['Longitud']}",  # Guardar solo las coordenadas
                'CONDUCTOR': '',
                'FECHA': row['Fecha']
            }
            registros.append(registro)
        
        print(f"Número total de registros: {len(registros)}")
        return registros

    # Ubicar
    def infracUbicarHistorico(self, file1):
        # Leer el archivo de Excel y obtener las hojas
        df = pd.read_excel(file1, skiprows=2).iloc[:-1]  # Ignorar las primeras 4 filas y la última fila

        # Extraer la placa del vehículo de la celda B1
        placa = "JYT620"  # Reemplazar con la extracción correcta si se requiere

        # Convertir la columna 'Comienzo' a datetime y extraer solo la fecha y hora
        df['Comienzo'] = pd.to_datetime(df['Comienzo'], dayfirst= True)
        df['Fecha'] = df['Comienzo'].dt.strftime('%d/%m/%Y %H:%M:%S')
        
        # Convertir la columna 'Duración' a segundos
        def convert_duration_to_seconds(duration):
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

        df['Tiempo de Exceso'] = df['Duración'].apply(convert_duration_to_seconds)
        
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
                'RUTA DE EXCESO': row['Posición'],
                'FECHA': row['Fecha']
            })

        # Imprimir el número de registros
        print(f'Número de registros: {len(registros)}')
        
        # Retornar el diccionario de registros
        return registros

    # Securitrac 
    def infracSecuritracHistorico(self, file):
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

    # Wialon
    def infracWialonHistorico(self, file1):
        xls = pd.ExcelFile(file1)

        # Leer la hoja 'Statistics' para obtener la placa
        df_stats = pd.read_excel(xls, 'Statistics')
        placa = df_stats.columns[1]  # Celda B1 en la hoja "Statistics"

        # Leer la hoja 'Excesos de velocidad' para obtener la información necesaria
        df_excesos = pd.read_excel(xls, 'Excesos de velocidad')

        # Convertir la columna 'Comienzo' a datetime y extraer solo la fecha y hora
        df_excesos['Comienzo'] = pd.to_datetime(df_excesos['Comienzo'])
        df_excesos['Fecha'] = df_excesos['Comienzo'].dt.strftime('%d/%m/%Y %H:%M:%S')

        # Convertir la columna 'Duración' a segundos
        def convert_duration_to_seconds(duration_str):
            parts = duration_str.split(':')
            if len(parts) == 3:
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
            elif len(parts) == 2:
                return int(parts[0]) * 60 + int(parts[1])
            else:
                return int(parts[0])

        df_excesos['Duración en segundos'] = df_excesos['Duración'].apply(convert_duration_to_seconds)

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
        return registros
    
    # Ejecutar todos
    def infracTodos(self, fileIturan, fileMDVR, fileUbicar, fileWialon1, fileWialon2, fileWialon3, fileSecuritrac):
        # Ejecutar cada función infrac y obtener los resultados
        registros_ituran = self.infracIturanHistorico(fileIturan)
        registros_mdvr = self.infracMDVRHistorico(fileMDVR)
        registros_ubicar = self.infracUbicarHistorico(fileUbicar)
        registros_wialon = self.infracWialonHistorico(fileWialon1)
        registros_wialon2 = self.infracWialonHistorico(fileWialon2)
        registros_wialon3 = self.infracWialonHistorico(fileWialon3)
        registros_securitrac = self.infracSecuritracHistorico(fileSecuritrac)

        # Combinar todos los resultados en una sola lista
        todos_registros = (
            registros_ituran + registros_mdvr + registros_ubicar + registros_wialon + registros_wialon2 + registros_wialon3 + registros_securitrac
        )

        #print(f"Total de registros combinados: {len(todos_registros)}")
        return todos_registros