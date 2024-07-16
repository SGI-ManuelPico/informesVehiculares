import pandas as pd
import openpyxl
import re
import xlrd
import os
from openpyxl import load_workbook
from datetime import datetime 
from openpyxl.utils.dataframe import dataframe_to_rows

from persistence.extraerExcel import ExtraerExcel

class ModificarExcel():

    def __init__(self):
        self.extraer = ExtraerExcel()

    # Crear el archivo Excel seguimiento.xlsx con los datos extraídos. Si el archivo ya existe, simplemente lo actualiza con los datos nuevos.

    def crear_excel3(self, mdvr_file1, mdvr_file2, ituran_file, ituran_file2, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2, output_file='seguimiento3.xlsx'):
        # Ejecutar todas las extracciones
        nuevos_datos = self.extraer.ejecutarTodasExtraccionesExcel(mdvr_file1, mdvr_file2, ituran_file, ituran_file2, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2)
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

    # Actualiza la hoja Infractores de la hoja de Excel. (Todavía falta testear esta función con otros archivos)
    def actualizarInfractores(self, file_seguimiento, file_Ituran, file_MDVR, file_Ubicar, file_Wialon, file_Securitrac):
        # Obtener todas las infracciones combinadas
        todos_registros = self.extraer.infracTodos(file_Ituran, file_MDVR, file_Ubicar, file_Wialon, file_Securitrac)
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

    # Actualizar hoja de Odómetro en el Excel
    def actualizarOdom(self, file_seguimiento, file_ituran, file_ubicar):
        # Obtener todos los odómetros combinados
        todos_registros = self.extraer.OdomIturan(file_ituran) + self.extraer.odomUbicar(file_ubicar)
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

# mdvr_file1 = r"C:\Users\SGI SAS\Downloads\general_information_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720804327.xls"
# mdvr_file2 = r"C:\Users\SGI SAS\Downloads\stops_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720804333.xlsx"
# archivoIturan1 = r"C:\Users\SGI SAS\Downloads\report.csv"
# archivoIturan2 = r"C:\Users\SGI SAS\Downloads\report(1).csv"
# securitrac_file = r"C:\Users\SGI SAS\Downloads\exported-excel.xls"
# wialon_file1 = r"C:\Users\SGI SAS\Downloads\LPN816_INFORME_GENERAL_TM_V1.0_2024-07-12_16-30-46.xlsx"
# wialon_file2 = r"C:\Users\SGI SAS\Downloads\LPN821_INFORME_GENERAL_TM_V1.0_2024-07-12_16-30-57.xlsx"
# wialon_file3 = r"C:\Users\SGI SAS\Downloads\JTV645_INFORME_GENERAL_TM_V1.0_2024-07-12_16-30-28.xlsx"
# ubicar_file1 = r"C:\Users\SGI SAS\Downloads\general_information_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720803683.xlsx"
# ubicar_file2 = r"C:\Users\SGI SAS\Downloads\stops_report_2024_07_11_00_00_00_2024_07_12_00_00_00_1720803694.xlsx"
# ubicom_file1 = r"C:\Users\SGI SAS\Downloads\ReporteDiario.xls"
# ubicom_file2 = r"C:\Users\SGI SAS\Downloads\Estacionados.xls"



# crear_excel(mdvr_file1, mdvr_file2, archivoIturan1, archivoIturan2, securitrac_file, wialon_file1, wialon_file2, wialon_file3, ubicar_file1, ubicar_file2, ubicom_file1, ubicom_file2, output_file='seguimiento.xlsx')

