import pandas as pd
import openpyxl
import re
import xlrd
import os
from openpyxl import load_workbook
from datetime import datetime 
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
import locale
import win32com.client as win32

from asdfg import Excelee


class Extracciones:
    def __init__(self):
        pass

        
    def crear_excel(self, mdvr_file1, mdvr_file2, ituran_file, ituran_file2, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2, output_file):
        # Ejecutar todas las extracciones
        nuevos_datos = Excelee().ejecutar_todas_extracciones(mdvr_file1, mdvr_file2, ituran_file, ituran_file2, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2)
        # Convertir la lista de nuevos datos a DataFrame
        df_nuevos = pd.DataFrame(nuevos_datos)

        if not os.path.exists(output_file):
            # Si el archivo no existe, crear el DataFrame inicial con el formato deseado
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

            for fecha in fechas:
                mes_dia = fecha.strftime('%d/%m')
                df_formato[mes_dia] = ''

            # Guardar el DataFrame en un archivo Excel
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_formato.to_excel(writer, sheet_name='Seguimiento', index=False)

        else:
            # Leer el archivo existente
            book = load_workbook(output_file)
            if 'Seguimiento' in book.sheetnames:
                sheet = book['Seguimiento']
                df_existente = pd.read_excel(output_file, sheet_name='Seguimiento')
            else:
                sheet = book.create_sheet('Seguimiento')
                df_existente = pd.DataFrame()

            # Rellenar los datos en el DataFrame con el formato deseado
            for _, row in df_nuevos.iterrows():
                fecha = pd.to_datetime(row['fecha'], format='%d/%m/%Y')
                dia = fecha.strftime('%d/%m')
                placa = row['placa']

                df_existente.loc[(df_existente['PLACA'] == placa) & (df_existente['SEGUIMIENTO'] == 'Nº Excesos'), dia] = row['num_excesos']
                df_existente.loc[(df_existente['PLACA'] == placa) & (df_existente['SEGUIMIENTO'] == 'Nº Desplazamiento'), dia] = row['num_desplazamientos']
                df_existente.loc[(df_existente['PLACA'] == placa) & (df_existente['SEGUIMIENTO'] == 'Día Trabajado'), dia] = row['dia_trabajado']
                df_existente.loc[(df_existente['PLACA'] == placa) & (df_existente['SEGUIMIENTO'] == 'Preoperacional'), dia] = row['preoperacional']
                df_existente.loc[(df_existente['PLACA'] == placa) & (df_existente['SEGUIMIENTO'] == 'Km recorridos'), dia] = row['km_recorridos']

            # Escribir los datos actualizados en la hoja 'seguimiento'
            for r_idx, row in enumerate(dataframe_to_rows(df_existente, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    sheet.cell(row=r_idx, column=c_idx, value=value)

            # Guardar el archivo Excel
            book.save(output_file)
            return df_existente

    
    def actualizarInfractores(self, file_seguimiento, file_Ituran, file_MDVR, file_Ubicar, fileWialon1, fileWialon2, fileWialon3, file_Securitrac):
        # Obtener todas las infracciones combinadas
        todos_registros = Excelee().infracTodos(file_Ituran, file_MDVR, file_Ubicar, fileWialon1, fileWialon2, fileWialon3, file_Securitrac)
        df_infractores = pd.DataFrame(todos_registros)

        # Convertir la columna 'FECHA' a datetime y luego a string con el formato correcto
        df_infractores['FECHA'] = pd.to_datetime(df_infractores['FECHA'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y %H:%M:%S')

        try:
            # Cargar el archivo existente y añadir una nueva hoja
            with pd.ExcelWriter(file_seguimiento, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
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

            print(f"Agregado como hoja 'Infractores' en {file_seguimiento}")

        except Exception as e:
            print(f"Error al actualizar el archivo Excel: {e}")


    def actualizarOdom(self, file_seguimiento, file_ituran, file_ubicar):
        # Obtener todos los odómetros combinados
        todos_registros = Excelee().OdomIturan(file_ituran) + Excelee().odomUbicar(file_ubicar)
        df_odometros = pd.DataFrame(todos_registros)

        # Cargar el archivo existente y añadir una nueva hoja, sobreescribiendo si ya existe
        with pd.ExcelWriter(file_seguimiento, engine='openpyxl', mode='a') as writer:
            # Verificar si la hoja ya existe
            if 'Odómetro' in writer.book.sheetnames:
                # Eliminar la hoja existente
                std = writer.book['Odómetro']
                writer.book.remove(std)
            # Escribir el DataFrame en una nueva hoja llamada 'Odometro'
            df_odometros.to_excel(writer, sheet_name='Odómetro', index=False)


    def actualizarIndicadoresTotales(self, df_diario, file_seguimiento):
        # Convertir la columna 'FECHA' a datetime y luego formatear para quitar la hora
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'], format='%Y-%m-%d')

        # Escribir el DataFrame en la hoja 'Indicadores Totales', reemplazando si existe
        with pd.ExcelWriter(file_seguimiento, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_diario.to_excel(writer, sheet_name='Indicadores Totales', index=False)

        print(f"Agregado como hoja 'Indicadores Totales' en {file_seguimiento}")

    
    def actualizarIndicadores(self, df_diario, df_exist, file_seguimiento):


        # Crear df_diario y df_hist
    

        # Calcular los cuatro indicadores
        df_EJL = Excelee().calcular_EJL(df_diario)
        df_GVE = Excelee().calcular_GVE(df_diario, df_exist)
        df_ELVL = Excelee().calcular_ELVL(df_diario)
        df_IDP = Excelee().calcular_IDP(df_diario)

        # Crear una lista de DataFrames
        dfs = [df_EJL, df_GVE, df_ELVL, df_IDP]

        # Crear el archivo si no existe
        if not os.path.exists(file_seguimiento):
            with pd.ExcelWriter(file_seguimiento, engine='xlsxwriter') as writer:
                # Crea un archivo de Excel vacío
                pd.DataFrame().to_excel(writer)

        # Cargar el archivo existente
        book = load_workbook(file_seguimiento)

        # Crear o seleccionar la hoja de trabajo 'Indicadores'
        if 'Indicadores' in book.sheetnames:
            del book['Indicadores']
        sheet = book.create_sheet('Indicadores')

        # Escribir cada DataFrame en la ubicación especificada con un espacio de 1 fila entre ellos
        start_row = 1
        for df in dfs:
            df = df.reset_index()
            df.rename(columns={'index': 'Indicador'}, inplace=True)
            
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row):
                for c_idx, value in enumerate(row):
                    cell = sheet.cell(row=r_idx, column=c_idx + 1, value=value)
                    if r_idx == start_row or c_idx == 0:  # Formato a los encabezados de los meses y la columna 'Indicador'
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            
            start_row += len(df) + 2  # Incrementar la fila inicial para el siguiente DataFrame (1 fila de espacio + 1 fila de encabezado)

        # Guardar el archivo
        book.save(file_seguimiento)


    def dfDiario(df_exist):
        sumas_diario = {
            'FECHA': [], 
            'KILOMETROS RECORRIDOS': [], 
            'EXCESOS VELOCIDAD': [], 
            'DESPLAZAMIENTOS': [], 
            'DÍA TRABAJADO': [], 
            'PREOPERACIONAL': []
        }

        date_columns = df_exist.columns[2:]

        
        for date in date_columns:
            sumas_diario['FECHA'].append(date)
            sumas_diario['KILOMETROS RECORRIDOS'].append(df_exist[df_exist['SEGUIMIENTO'] == 'Km recorridos'][date].sum())
            sumas_diario['EXCESOS VELOCIDAD'].append(df_exist[df_exist['SEGUIMIENTO'] == 'Nº Excesos'][date].sum())
            sumas_diario['DESPLAZAMIENTOS'].append(df_exist[df_exist['SEGUIMIENTO'] == 'Nº Desplazamiento'][date].sum())
            sumas_diario['DÍA TRABAJADO'].append(df_exist[df_exist['SEGUIMIENTO'] == 'Día Trabajado'][date].sum())
            sumas_diario['PREOPERACIONAL'].append(df_exist[df_exist['SEGUIMIENTO'] == 'Preoperacional'][date].sum())

        df_diario = pd.DataFrame(sumas_diario)
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'] + '/2024', format='%d/%m/%Y').dt.strftime('%Y-%m-%d')  # Toca ajustar el año según lo necesitemos.
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'])

        return df_diario


archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"



df_exist = Extracciones().crear_excel
Extracciones().actualizarInfractores
Extracciones().actualizarOdom
df_diario = Extracciones().dfDiario(df_exist)
Extracciones().actualizarIndicadoresTotales
Extracciones().actualizarIndicadores