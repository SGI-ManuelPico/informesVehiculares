import pandas as pd
from persistence.scriptMySQL import ActualizadorSQL
import pandas as pd
import sqlalchemy
from sqlalchemy import create_engine, text, Table, update
from sqlalchemy.orm import sessionmaker
from persistence.archivoExcel import FuncionalidadExcel
from db.conexionDB import conexionDB



class FuncionalidadSQL:
    def __init__(self):
        pass

    def actualizarSeguimientoSQL(self, file_ituran, file_ituran2, file_MDVR1, file_MDVR2, file_Ubicar1, file_Ubicar2, file_Ubicom1, file_Ubicom2, file_Securitrac, file_Wialon1, file_Wialon2, file_Wialon3):

    # Ejecutar cada función de extracción con los archivos proporcionados
        datos_mdvr = ActualizadorSQL().sqlMDVR(file_MDVR1, file_MDVR2)
        datos_ituran = ActualizadorSQL().sqlIturan(file_ituran, file_ituran2)
        datos_securitrac = ActualizadorSQL().sqlSecuritrac(file_Securitrac)
        datos_wialon = ActualizadorSQL().sqlWialon(file_Wialon1, file_Wialon2, file_Wialon3)
        datos_ubicar = ActualizadorSQL().sqlUbicar(file_Ubicar1, file_Ubicar2)
        datos_ubicom = ActualizadorSQL().sqlUbicom(file_Ubicom1, file_Ubicom2)

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

        user = 'root'
        password = 'Gatitos24'  
        host = '127.0.0.1'
        port = '3306'
        schema = 'vehiculos'

        engine = create_engine(f'mysql+mysqlconnector://{user}:{password}@{host}:{port}/{schema}')

        # Insertar los datos en la tabla seguimiento
        df_final.to_sql('seguimiento', con=engine, if_exists='append', index=False)

        print("Datos insertados correctamente en la tabla seguimiento.")


    def actualizarInfractoresSQL(self, fileIturan, fileMDVR, fileUbicar, fileWialon, fileWialon2, fileWialon3, fileSecuritrac):

        # Obtener todas las infracciones combinadas
        registros_ituran = FuncionalidadExcel().infracIturan(fileIturan)
        registros_mdvr = FuncionalidadExcel().infracMDVR(fileMDVR)
        registros_ubicar = FuncionalidadExcel().infracUbicar(fileUbicar)
        registros_wialon = FuncionalidadExcel().infracWialon(fileWialon)
        registros_wialon2 = FuncionalidadExcel().infracWialon(fileWialon2)
        registros_wialon3 = FuncionalidadExcel().infracWialon(fileWialon3)
        registros_securitrac = FuncionalidadExcel().infracSecuritrac(fileSecuritrac)

        # Combinar todos los resultados en una sola lista
        todos_registros = (
            registros_ituran + registros_mdvr + registros_ubicar + registros_wialon + registros_wialon2 + registros_wialon3 +registros_securitrac
        )


        df_infractores = pd.DataFrame(todos_registros)

        # Convertir la columna 'FECHA' a datetime y luego a string con el formato correcto
        df_infractores['FECHA'] = pd.to_datetime(df_infractores['FECHA'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')

        df_infractores = df_infractores[(df_infractores['VELOCIDAD MÁXIMA'] > 80) & (df_infractores['TIEMPO DE EXCESO'] > 20)]

        # Renombrar las columnas a camelCase
        df_infractores.columns = [
            'placa', 
            'tiempoDeExceso', 
            'duracionEnKmDeExceso', 
            'velocidadMaxima', 
            'proyecto', 
            'rutaDeExceso', 
            'conductor', 
            'fecha'
        ]

        # Conexión a la base de datos MySQL
        engine = sqlalchemy.create_engine('mysql+mysqlconnector://root:Gatitos24@127.0.0.1:3306/vehiculos')

        try:
            # Leer los datos existentes en la tabla 'infractores'
            df_existente = pd.read_sql_table('infractores', con=engine)

            # Concatenar los datos existentes con los nuevos datos
            df_final = pd.concat([df_existente, df_infractores], ignore_index=True)

            # Eliminar duplicados opcionalmente
            df_final = df_final.drop_duplicates(subset=['placa', 'fecha', 'tiempoDeExceso', 'duracionEnKmDeExceso', 'velocidadMaxima', 'proyecto', 'rutaDeExceso', 'conductor'])

        except ValueError:  # Si la tabla no existe aún, usa los nuevos datos directamente
            df_final = df_infractores

        # Exportar el DataFrame combinado a la tabla 'infractores'
        df_final.to_sql(name='infractores', con=engine, if_exists='replace', index=False)

        print("Datos actualizados en la tabla 'infractores'")
        return df_final
    

    def actualizarKilometraje(self, file_ituran, file_ubicar):
        todos_registros = FuncionalidadExcel().OdomIturan(file_ituran) + FuncionalidadExcel().odomUbicar(file_ubicar)
        df_odometro = pd.DataFrame(todos_registros)


        user = 'root'
        password = 'Gatitos24'
        host = '127.0.0.1'
        port = '3306'
        schema = 'vehiculos'
        tabla = 'carro'

        # Crear la cadena de conexión usando las variables
        engine = create_engine(f'mysql+mysqlconnector://{user}:{password}@{host}:{port}/{schema}')

        Session = sessionmaker(bind=engine)
        session = Session()

        consulta_sql = text(f'''
        UPDATE {tabla}
        SET {'Kilometraje'} = :nuevo_valor
        WHERE {'Placas'} = :primary_key_value
        ''')
        for index, placa in df_odometro.iterrows():

            kilometraje = placa['KILOMETRAJE']
            placa1 = placa['PLACA']

            session.execute(consulta_sql, {'nuevo_valor': kilometraje, 'primary_key_value': placa1})
            session.commit()
        session.close() 

    ## Esta función recibe como argumentos el df que sale de la función fueraLaboralTodos y db_conn es una instancia de la clase ConexionDB.

    def sqlFueraLaboral(self, df):
    
        # Cambiar la fecha al formato de curdate()

        df['fecha'] = pd.to_datetime(df['fecha'], format='%d/%m/%Y %H:%M').dt.strftime('%Y-%m-%d %H:%M:%S')

        conexionBase = conexionDB().establecerConexion()
        if conexionBase:
            cursor = conexionBase.cursor()
        else:
            print("Error.")

        # Convertir el DataFrame en una lista de tuplas
        data = [tuple(row) for row in df.values]
        
        # Definir la consulta SQL para insertar datos
        insert_query = """
        INSERT INTO fueraLaboral (placa, fecha)
        VALUES (%s, %s)
        """

        # Ejecutar la consulta para cada fila de datos
        cursor.executemany(insert_query, data)
        
        # Confirmar los cambios
        cursor.commit()

        conexionDB().cerrarConexion()


        
        
