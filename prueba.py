from db.conexionDB import conexionDB
from forms.rpaCompleto import RPA
from forms.wialonForm import DatosWialon
from util.tratadoArchivos import TratadorArchivos
from forms.MDVRForm import DatosMDVR
from persistence.archivoExcel import FuncionalidadExcel
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


# wialon = DatosWialon()

# mdvr = DatosMDVR()

# mdvr.rpaMDVR()

functionalidad = FuncionalidadExcel()

datos = functionalidad.extraerMDVR(r'C:\Users\SGI SAS\Documents\GitHub\SGI\outputMDVR\general_information_report_2024_08_01_00_00_00_2024_08_02_00_00_00_1722522962.xls', r'C:\Users\SGI SAS\Documents\GitHub\SGI\outputMDVR\overspeeds_report_2024_08_01_00_00_00_2024_08_02_00_00_00_1722522970.xlsx' )

print(datos)
