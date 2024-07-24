import pandas as pd
from datetime import datetime
from db.conexionDB import conexionDB 
import mysql.connector

class EstadoPlataforma:
    def __init__(self):
        self.db = conexionDB()

    def actualizar_estado(self, plataforma, estado):
        connection = None
        cursor = None
        try:
            connection = self.db.establecerConexion()
            if connection is None:
                return
            cursor = connection.cursor()
            update_query = f"UPDATE estadoPlataforma SET estado = '{estado}' WHERE plataforma = '{plataforma}'"
            cursor.execute(update_query)
            connection.commit()
        except mysql.connector.Error as err:
            print(f"Error: {err}")
        finally:
            if cursor:
                cursor.close()
            self.db.cerrarConexion()

    def verificar_estado(self):
        connection = None
        cursor = None
        try:
            connection = self.db.establecerConexion()
            if connection is None:
                return []
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
            self.db.cerrarConexion()

    def log_error(self, plataforma):
        connection = None
        cursor = None
        try:
            connection = self.db.establecerConexion()
            if connection is None:
                return
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
            self.db.cerrarConexion()

    def reset_estados(self):
        connection = None
        cursor = None
        try:
            connection = self.db.establecerConexion()
            if connection is None:
                return
            cursor = connection.cursor()
            reset_query = "UPDATE estadoPlataforma SET estado = 'no ejecutado'"
            cursor.execute(reset_query)
            connection.commit()
        except mysql.connector.Error as e:
            print(f"Error: {e}")
        finally:
            if cursor:
                cursor.close()
            self.db.cerrarConexion()

    def update_error(self):
        pass  
