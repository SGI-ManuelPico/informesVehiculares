# from archivoExcel import extraerIturan
import pandas as pd
import datetime as datetime
# Extraer los datos de los informes de Ituran.

def extraerIturan(file1, file2):

    try:
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
    
    # Si no se pueden sacar los archivos de la plataforma por alguna razón:

    except Exception as e: 
        print('Archivos incorrectos o faltantes ituran')
        return []

# Extraer los datos de los informes de MDVR.
