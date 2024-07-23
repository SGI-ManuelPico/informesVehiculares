from db.conexionDB import conexionDB

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
        # PARA CADA PLATAFORMA VERIFICAR EN LA TABLA QUE SALE DE AQUÍ SI YA SE HIZO Y SI NO HACER EL RPA.

    def actualizarEstadoPlataforma(self, plataforma,estado):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"""UPDATE vehiculos.estadosvehiculares SET estado = "{estado}" WHERE plataforma = '{plataforma}'""")

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

    def crearTablaEstados(self):
        """
        Crea la tabla de estados para cada día.
        """
        
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute("DROP TABLE IF EXISTS vehiculos.estadosVehiculares")

        cursor.execute(f"""CREATE TABLE vehiculos.estadosVehiculares(
                        id int NOT NULL,
                        plataforma varchar(20),
                        estado varchar(20),
                        PRIMARY KEY (`plataforma`)
                        );""")

        consultaDentro = "insert into vehiculos.estadosVehiculares (id, plataforma, estado) values (%s, %s, %s);"
        consultaInserto = [('1', 'Ituran', 'No ejecutado'),
                            ('2', 'Securitrac', 'No ejecutado'),
                            ('3', 'MDVR', 'No ejecutado'),
                            ('4', 'Ubicar', 'No ejecutado'),
                            ('5', 'Ubicom', 'No ejecutado'),
                            ('6', 'Wialon', 'No ejecutado')]

        cursor.executemany(consultaDentro,consultaInserto)
        conexionBaseCorreos.commit()

        #Desconectar BD
        conexionDB().cerrarConexion()

