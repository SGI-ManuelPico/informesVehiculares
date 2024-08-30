from db.conexionDB import conexionDB
from datetime import datetime
import mysql.connector

class ConsultaImportante:
    def tablaCorreoPersonal(self):
        """
        Busca las tablas para los correos normales que se piden.
        """
        
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute("SELECT placa, tiempoDeExceso, velocidadMaxima, conductor FROM vehiculos.infractores where date(fecha) = curdate();")
        self.tablaExcesos = cursor.fetchall()
        cursor.execute("select * from vehiculos.correovehicular")
        self.tablaCorreos = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaExcesos, self.tablaCorreos

    def tablaCorreoPlataforma(self):
        """
        Busca las tablas para los correos que son notificados de fallas en plataformas.
        """
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute("select * from vehiculos.correovehicular")
        self.tablaCorreos2 = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaCorreos2

    def tablaEstadosPlataforma(self, plataforma):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"""select estado from vehiculos.estadosvehiculares where plataforma = '{plataforma}'""")
        self.tablaEstados = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaEstados

    def actualizarEstadoPlataforma(self, plataforma, estado):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"""UPDATE vehiculos.estadosvehiculares SET estado = "{estado}" WHERE (plataforma = '{plataforma}')""")

        conexionBaseCorreos.commit()

        #Desconectar BD
        conexionDB().cerrarConexion()
        
    def verificarEstadosFinales(self):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"""select plataforma, estado from vehiculos.estadosvehiculares""")
        self.tablaEstadosTotales = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaEstadosTotales

    def actualizarTablaEstados(self):
        """
        Crea la tabla de estados para cada día.
        """
        
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        plataformasVehiculares = ["Ituran", "Securitrac", "MDVR","Ubicar","Ubicom","Wialon"]
        for plataforma in plataformasVehiculares:
            cursor.execute(f"""UPDATE `vehiculos`.`estadosvehiculares` SET `estado` = 'No ejecutado' WHERE (`plataforma` = '{plataforma}');""")
            conexionBaseCorreos.commit()
        
        #Desconectar BD
        conexionDB().cerrarConexion()

    def registrarError(self, plataforma):

        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")
        
        fecha = datetime.now().strftime('%d/%m/%Y')
        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"INSERT INTO vehiculos.tablaerrores (plataforma, fecha, estado) VALUES ('{plataforma}', '{fecha}', 'error')")
        conexionBaseCorreos.commit()
        #Desconectar BD
        conexionDB().cerrarConexion()

    def tablaWialon(self):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")
        
        #Consulta de las placas que componen a Wialon.
        cursor.execute("select placa, plataforma from vehiculos.placasvehiculos where plataforma = 'Wialon'")
        self.placasPWialon = cursor.fetchall() #Obtener todos los resultados
        
        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.placasPWialon

    def tablaCorreoLaboral(self):
        """
        Busca la tabla para el correo de vehículos con desplazamientos fuera de horario laboral.
        """
        
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute("SELECT placa, fecha, conductor FROM vehiculos.fueraLaboral where date(fecha) like curdate();")
        self.tablaHorarios = cursor.fetchall()
        cursor.execute("SELECT placa, conductor FROM vehiculos.fueralaboral where date(fecha) like curdate();")
        self.tablaPuntos = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaHorarios, self.tablaPuntos
    
    def actualizarEstadoError(self, error_id, status):
        conexionBaseErrores = conexionDB().establecerConexion()
        if conexionBaseErrores:
            cursor = conexionBaseErrores.cursor()
            try:
                update_query = f"UPDATE vehiculos.tablaerrores SET estado = '{status}' WHERE id = {error_id}"
                cursor.execute(update_query)
                conexionBaseErrores.commit()
            except mysql.connector.Error as err:
                print(f"Error: {err}")
            finally:
                cursor.close()
                conexionDB().cerrarConexion()
        else:
            print("Error al conectar a la base de datos.")

    def checkCamposError(self):
        conexionBaseErrores = conexionDB().establecerConexion()
        if conexionBaseErrores:
            cursor = conexionBaseErrores.cursor(dictionary=True)
            try:
                query = "SELECT id, plataforma, fecha FROM vehiculos.tablaerrores WHERE estado = 'error'"
                cursor.execute(query)
                error_entries = cursor.fetchall()
                return error_entries
            except mysql.connector.Error as err:
                print(f"Error: {err}")
                return []
            finally:
                cursor.close()
                conexionDB().cerrarConexion()
        else:
            print("Error al conectar a la base de datos.")
            return []