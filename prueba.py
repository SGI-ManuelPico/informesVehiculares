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