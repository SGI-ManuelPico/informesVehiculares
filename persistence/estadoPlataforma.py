import pandas as pd
import mysql.connector
from datetime import datetime

#### Acá puse las conexiones a la base de datos en todas las funciones, pero supongo que con la clase de conexión a DB podemos arreglar esa redundancia.

def actualizarEstado(plataforma, estado):
    try:
        # Conectarse a la base de datos.
        connection = mysql.connector.connect(
            host='host',
            user='user',
            password='password',
            database='database'
        )

        # Actualizar el estado de la ejecución del RPA

        cursor = connection.cursor()
        update_query = f"UPDATE estadoPlataforma SET estado = '{estado}' WHERE plataforma = '{plataforma}'"
        cursor.execute(update_query)
        connection.commit()
    except mysql.connector.Error as err:
        print(f"Error: {err}")
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()


def verificarEstado():
    try:
        # Conexión a la base de datos
        connection = mysql.connector.connect(
            host='host',
            user='user',
            password='password',
            database='database'
        )
        cursor = connection.cursor()
        query = "SELECT plataforma, estado FROM estadoPlataforma"
        cursor.execute(query)
        results = cursor.fetchall()
        return results
    except mysql.connector.Error as e:
        print(f"Error: {e}")
        return []
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()   


### Acá nos toca ver como hacemos para que esto se actualice cuando manden los archivos faltantes. 

def logError(plataforma):
    try:
        # Conexión a la base de datos
        connection = mysql.connector.connect(
            host='your_host',
            user='your_user',
            password='your_password',
            database='your_database'
        )
        cursor = connection.cursor()
        fecha = datetime.now().strftime('%d/%m/%Y')
        insert_query = f"INSERT INTO error (plataforma, fecha, estado) VALUES ('{plataforma}', '{fecha}', 'error')"
        cursor.execute(insert_query)
        connection.commit()
    except mysql.connector.Error as e:
        print(f"Error: {e}")
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()


 ## Esta todavia no se como definirla porque toca ver como va a ser bien este proceso.
 
def updateEror():
    return None


def resetEstados():
    try:
        # Conexión a la base de datos.
        connection = mysql.connector.connect(
            host='your_host',
            user='your_user',
            password='your_password',
            database='your_database'
        )
        cursor = connection.cursor()
        reset_query = "UPDATE estadoPlataforma SET estado = 'no ejecutado'"
        cursor.execute(reset_query)
        connection.commit()
    except mysql.connector.Error as e:
        print(f"Error: {e}")
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()
