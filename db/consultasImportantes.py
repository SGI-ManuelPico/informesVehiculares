from db.conexionDB import conexionDB
from datetime import datetime

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
        cursor.execute("SELECT placa, tiempoDeExceso, velocidadMaxima, conductor FROM vehiculos.infractores where date(fecha) like curdate();")
        self.tablaExcesos = cursor.fetchall()
        cursor.execute("select * from vehiculos.plataformasVehiculares")
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
        cursor.execute("select * from vehiculos.plataformasVehiculares")
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
        cursor.execute(f"""select estado from vehiculos.estadosVehiculares where plataforma = '{plataforma}'""")
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
        cursor.execute(f"""UPDATE vehiculos.estadosvehiculares SET estado = "{estado}" WHERE plataforma = '{plataforma}'""")

        #Desconectar BD
        conexionDB().cerrarConexion()
        
    def verificarEstadosFinales(self):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"""select plataforma, estado from vehiculos.estadosVehiculares""")
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

        plataformasVehiculares = ["Ituran", "Securitac", "MDVR","Ubicar","Ubicom","Wialon"]
        for plataforma in plataformasVehiculares:
            cursor.execute(f"""UPDATE `vehiculos`.`estadosvehiculares` SET `estado` = 'No ejecutado' WHERE (`plataforma` = '{plataforma}');""")

        #Desconectar BD
        conexionDB().cerrarConexion()

    def registrarError(plataforma):

        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")
        
        fecha = datetime.now().strftime('%d/%m/%Y')
        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"INSERT INTO error (plataforma, fecha, estado) VALUES ('{plataforma}', '{fecha}', 'error')")

        #Desconectar BD
        conexionDB().cerrarConexion()

    def tablaWialon(self):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")
        
        #Consulta de las placas que componen a Wialon.
        cursor.execute("select placa, plataforma from vehiculos.placasVehiculos where plataforma = 'Wialon'")
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
        cursor.execute("SELECT * FROM vehiculos.fueraLaboral where date(fecha) like curdate();")
        self.tablaHorarios = cursor.fetchall()
        cursor.execute("SELECT placa FROM vehiculos.fueraLaboral where date(fecha) like curdate();")
        self.tablaPuntos = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaHorarios, self.tablaPuntos