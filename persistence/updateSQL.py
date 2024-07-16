import pandas as pd
import openpyxl
import re
import xlrd
from datetime import datetime
from tkinter import messagebox
import warnings
import numpy as np
import math

from persistence.extraerExcel import ExtraerExcel
from db.conexionDB import conexionDB

class ActualizarBD(conexionDB):
    """
    Esta clase sirve para crear todos los metodos que vayan a realizar consultas SQL
    referentes a la actualización de la base de datos de Vehiculos.
    """

    def __init__(self):
        super().__init__()
        self.extraer = ExtraerExcel()

    def actualizarSeguimientoSQL(self, file_ituran1, file_ituran2, file_MDVR1, file_MDVR2, file_Ubicar1, file_Ubicar2, file_Ubicom1, file_Ubicom2, file_Securitrac, file_Wialon1, file_Wialon2, file_Wialon3):

        df_seguimiento = self.extraer.ejecutarTodasExtraccionesSQL(
        file_ituran1,
        file_ituran2, 
        file_MDVR1, 
        file_MDVR2, 
        file_Ubicar1, 
        file_Ubicar2,
        file_Ubicom1, 
        file_Ubicom2,
        file_Securitrac,
        file_Wialon1,
        file_Wialon2, 
        file_Wialon3
        )

        # Suprimir todas las advertencias
        warnings.filterwarnings("ignore")
        #print(f"placa: {df_seguimiento.iloc[12][0]}\nkmRecorridos: {df_seguimiento.iloc[12][1]}\nnumDesplazamientos: {df_seguimiento.iloc[12][2]}\ndiaTrabajado: {df_seguimiento.iloc[12][3]}\npreoperacional: {df_seguimiento.iloc[12][4]}\nnumExcesos: {df_seguimiento.iloc[12][5]}\nproveedor: {df_seguimiento.iloc[12][6]}\nfecha: {df_seguimiento.iloc[12][7]}\n")

        #Conectar BD
        self.conexion = self.establecerConexion()
        if self.conexion:
            self.cursor = self.conexion.cursor()
        else:
            messagebox.showinfo("Error","Error al tomar los nombres de los empleados.")

        for i in range(len(df_seguimiento)):
            valores_convertidos = [int(x) if isinstance(x, np.integer) else float(x) if isinstance(x, np.floating) else x for x in df_seguimiento.iloc[i]]
            placa = valores_convertidos[0]
            kmRecorridos = valores_convertidos[1]
            numDesplazamientos = valores_convertidos[2]
            diaTrabajado = valores_convertidos[3]
            preoperacional = valores_convertidos[4]
            numExcesos = valores_convertidos[5]
            proveedor = valores_convertidos[6]
            fecha = valores_convertidos[7]
            #Consulta
            valores = (placa, kmRecorridos, numDesplazamientos, diaTrabajado, preoperacional, numExcesos, proveedor, fecha)
            consulta = "INSERT INTO seguimiento (placa, kmRecorridos, numDesplazamientos, diaTrabajado, preoperacional, numExcesos, proveedor, fecha) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
            self.cursor.execute(consulta, valores)
            self.conexion.commit()

        # Verificar si la consulta se realizó correctamente
        if self.cursor.rowcount > 0:
            print("Datos insertados correctamente en la tabla seguimiento.")
        else:
            print("Error.") 
    
        #Desconectar BD
        self.cerrarConexion()


    # Actualizar infractores (Esto ya está en otra parte, me toca moverlo acá)

    def actualizarKilometraje(self, file_ituran, file_ubicar):
        todos_registros = self.extraer.OdomIturan(file_ituran) + self.extraer.odomUbicar(file_ubicar)
        df_odometro = pd.DataFrame(todos_registros)

        #Conectar BD
        self.conexion = self.establecerConexion()
        
        if self.conexion:
            self.cursor = self.conexion.cursor()
        else:
            messagebox.showinfo("Error","Error al tomar los nombres de los empleados.")
        
        #Consulta
        consulta = "UPDATE carro SET Kilometraje = %s WHERE Placas = %s"
        
        for index, fila in df_odometro.iterrows():

            kilometraje = fila['KILOMETRAJE']
            placa = fila['PLACA']

            self.cursor.execute(consulta, (kilometraje, placa))
            self.conexion.commit()

        #Desconectar BD
        self.cerrarConexion()

    # Esta función es la que llena la tabla de infractores en la base de datos 'vehiculos' con la información histórica. 
    ###### ESTO ES PARA ACTUALIZAR LAS BASES DE DATOS EN MySQL ########
    def actualizarInfractoresSQL(self, file_Ituran, file_MDVR, file_Ubicar, file_Wialon1, file_Wialon2, file_Wialon3, file_Securitrac): 
        # Obtener todas las infracciones combinadas
        todos_registros = self.extraer.infracTodos(file_Ituran, file_MDVR, file_Ubicar, file_Wialon1, file_Wialon2, file_Wialon3, file_Securitrac)
        df_infractores = pd.DataFrame(todos_registros)

        # Convertir la columna 'FECHA' a datetime y luego a string con el formato correcto
        df_infractores['FECHA'] = pd.to_datetime(df_infractores['FECHA'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y %H:%M:%S')

        #Conectar BD
        self.conexion = self.establecerConexion()
        
        if self.conexion:
            self.cursor = self.conexion.cursor()
        else:
            messagebox.showinfo("Error","Error al tomar los nombres de los empleados.")

        for i in range(len(df_infractores)):
            valores_convertidos = [int(x) if isinstance(x, np.integer) else float(x) if isinstance(x, np.floating) else x for x in df_infractores.iloc[i]]

            # Reemplaza los valores 'nan' con None
            valores_convertidos = [None if (isinstance(val, float) and math.isnan(val)) else val for val in valores_convertidos]

            placa = valores_convertidos[0]
            tiempo_exceso = valores_convertidos[1]
            ruta_exceso = valores_convertidos[5]
            kms_exceso = valores_convertidos[2]
            velocidad_exceso = valores_convertidos[3]
            proyecto = valores_convertidos[4]
            conductor = valores_convertidos[6]
            fecha_evento = valores_convertidos[7]

            if kms_exceso is not None:
                try:
                    kms_exceso = float(kms_exceso)
                except ValueError:
                    kms_exceso = None
            elif kms_exceso is None:
                kms_exceso = None

            if proyecto is None:
                proyecto = ""

            if conductor is None:
                conductor = ""

            #Consulta
            valores = (placa, tiempo_exceso, ruta_exceso, kms_exceso, velocidad_exceso, proyecto, conductor, fecha_evento)
            consulta = "INSERT INTO infractor (placa, tiempo_exceso, ruta_exceso, kms_exceso, velocidad_exceso, proyecto, conductor, fecha_evento) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
            self.cursor.execute(consulta, valores)
            self.conexion.commit()

        # Verificar si la consulta se realizó correctamente
        if self.cursor.rowcount > 0:
            print("Datos insertados correctamente en la tabla seguimiento.")
        else:
            print("Error.")

        #Desconectar BD
        self.cerrarConexion()
    

