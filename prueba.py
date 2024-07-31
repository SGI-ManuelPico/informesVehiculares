from db.conexionDB import conexionDB
from forms.rpaCompleto import RPA
from forms.wialonForm import DatosWialon
from util.tratadoArchivos import TratadorArchivos
# print("Script started")  # Debugging print
# db = conexionDB()
# print("conexionDB instance created")  # Debugging print

# connection = db.establecerConexion()
# print("establish connection called")  # Debugging print

# if connection:
#     print("Connection successful")
#     db.cerrarConexion()
#     print("Connection closed")  # Debugging print
# else:
#     print("Connection failed")


wialon = DatosWialon()

TratadorArchivos().eliminarArchivosPlataforma("Wialon")