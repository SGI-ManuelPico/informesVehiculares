import pandas as pd

class Excelee:
    def __init__(self):
        pass


    def OdomIturan(self,file):
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

    def odomUbicar(self,file):
        # Leer el archivo de Excel
        df = pd.read_excel(file)  

        # Extraer la placa del vehículo de la celda B1, si agregan otro carro a esta plataforma toca cambiar como se extrae esto.
        placa = 'JYT620'

        # Extraer el odómetro de la celda correspondiente
        odometro = df.iloc[11, 2] 

        # Crear el diccionario con el formato requerido
        registro = {
            'PLACA': placa,
            'KILOMETRAJE': float(odometro.split()[0].replace(',', ''))
        }

        return [registro]
