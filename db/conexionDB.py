import mysql.connector
from tkinter import messagebox

class conexionDB:
    def __init__(self):
        self.host = 'localhost'
        self.user = 'root'
        self.password = '12345678'
        self.database = 'vehiculos'
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