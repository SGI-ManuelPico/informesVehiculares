import os, glob
import pandas as pd
lugarDescargasWialon = os.getcwd() + "\\outputWialon"
archivos = glob.glob(os.path.join(lugarDescargasWialon, '*.xlsx'))
archivoWialon1=archivoWialon2=archivoWialon3 = str()

import mysql.connector
from tkinter import messagebox

class conexionDB:
    def __init__(self):
        self.host = '127.0.0.1'
        self.user = 'root'
        self.password = 'Gatitos24'
        self.database = ''
        self.conexion = None
        
    def establecerConexion(self):
        try:
            self.conexion = mysql.connector.connect(
                host = self.host,
                user = self.user,
                password = self.password,
                database = self.database
            )
            return self.conexion
        
        except mysql.connector.Error as e:
            messagebox.showerror(message=f'Error de conexión: {e}', title='Mensaje')
            return None
        
    def cerrarConexion(self):
        if self.conexion:
            self.conexion.close()

# Tabla del correo.
conexionBaseCorreos = conexionDB().establecerConexion()
if conexionBaseCorreos:
    cursor = conexionBaseCorreos.cursor()
else:
    print("Error.")

#Consulta de las placas que componen a Wialon.
cursor.execute("select placa, plataforma from vehiculos.placasVehiculos where plataforma = 'Wialon'")
placasPWialon = cursor.fetchall() #Obtener todos los resultados

#Desconectar BD
conexionDB().cerrarConexion()

##########
placasPWialon = pd.DataFrame(placasPWialon, columns=['Placa', 'plataforma'])
placasWialon = placasPWialon['Placa'].tolist()


print(archivos)
for archivo in archivos:
    for placa in placasWialon:
        if placa in archivo and placa == placasWialon[0]:
            archivoWialon1 += archivo
        else:
            archivoWialon1
        if placa in archivo and placa == placasWialon[1]:
            archivoWialon2 += archivo
        else:
            archivoWialon2
        if placa in archivo and placa == placasWialon[2]:
            archivoWialon3 += archivo
        else:
            archivoWialon3

print(archivoWialon1, archivoWialon2, archivoWialon3)


def extraerWialon(file_path1, file_path2, file_path3):

    try:

        datos_extraidos = []
        
        for file_path in [file_path1, file_path2, file_path3]:
            xl = pd.ExcelFile(file_path)
            
            # Extraer placa y fecha siempre
            if 'Statistics' in xl.sheet_names:
                statistics_df = xl.parse('Statistics', header=None)
                placa = statistics_df.iloc[0, 1]  # Celda B1
                fecha = statistics_df.iloc[1, 1].split()[0].replace('.', '/')  # Celda B2
                fecha_formateada = pd.to_datetime(fecha).strftime('%d/%m/%Y')
                km_recorridos = int(statistics_df.iloc[7, 1])  # Celda B8, quitando 'km'
                dia_trabajado = 1 if km_recorridos > 0 else 0
                preoperacional = 1 if dia_trabajado == 1 else 0
            
                datos = {
                    'placa': placa,
                    'fecha': fecha_formateada,
                    'km_recorridos': km_recorridos,
                    'dia_trabajado': dia_trabajado,
                    'preoperacional': preoperacional
                }

            else:
                datos = {
                    'placa': placa,
                    'fecha': fecha_formateada,
                    'km_recorridos': 0,
                    'dia_trabajado': 0,
                    'preoperacional': 0,
                }
            
            # Verificar si el archivo tiene datos
            if 'Excesos de velocidad' in xl.sheet_names:
                excesos_df = xl.parse('Excesos de velocidad', header=None)
                num_excesos = len(excesos_df) - 1  # Descontar la fila de encabezado
                datos.update({'num_excesos': num_excesos})
                
            else:
                datos.update({'num_excesos': 0})

            
            # Extraer número de desplazamientos
            if 'Cronología' in xl.sheet_names:
                crono = xl.parse('Cronología')
                desplazamientos = 0
                for x in crono['Tipo'].to_list():
                    if x == 'Trip':
                        desplazamientos += 1
                
                datos.update({'num_desplazamientos': desplazamientos})

            else:
                datos.update({'num_desplazamientos': 0})
                
            
            datos_extraidos.append(datos)
        
        return datos_extraidos
    
    except Exception as e: 
        print('Archivos incorrectos o faltantes WIALON')
        return []


datos = extraerWialon(archivoWialon1, archivoWialon2, archivoWialon3)

print(datos)