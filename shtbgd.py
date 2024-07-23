from db.conexionDB import conexionDB
import pandas as pd

class aaa:

    def verificarEstadosFinales(self):
        conexionBaseCorreos = conexionDB().establecerConexion()
        if conexionBaseCorreos:
            cursor = conexionBaseCorreos.cursor()
        else:
            print("Error.")

        #Consulta de los correos necesarios para el correo.
        cursor.execute(f"""select estado from vehiculos.estadosVehiculares""")
        self.tablaEstadosTotales = cursor.fetchall()

        #Desconectar BD
        conexionDB().cerrarConexion()

        return self.tablaEstadosTotales
    


a = aaa().verificarEstadosFinales()
a = pd.DataFrame(a)

print(a)

# for e in a:
#     print(e)
#     if e == "('No ejecutado',)":
#         print(e)
#     else:
#         print("ASAA")