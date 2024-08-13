from db.conexionDB import conexionDB
from forms.rpaCompleto import RPA
from forms.wialonForm import DatosWialon
from util.tratadoArchivos import TratadorArchivos
from forms.MDVRForm import DatosMDVR
from forms.securitracForm import DatosSecuritrac
from persistence.archivoExcel import FuncionalidadExcel
from persistence.extraerExcel import Extracciones
from db.consultasImportantes import ConsultaImportante
import xlrd
import openpyxl
import pandas as pd



a,b=ConsultaImportante().tablaCorreoPersonal()
print(a)

# consulta = ConsultaImportante()
# wialon = DatosWialon()
# mdvr = DatosMDVR()
# secu = DatosSecuritrac()

# rpa = RPA()
# functionalidad = FuncionalidadExcel()

# extracciones = Extracciones()


# file_ubicar1 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputUbicar\general_information_report_2024_08_01_00_00_00_2024_08_02_00_00_00_1722540202.xlsx"
# file_ubicar2 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputUbicar\overspeeds_report_2024_08_01_00_00_00_2024_08_02_00_00_00_1722540205.xlsx"
# file_mdvr1 = r'C:\Users\SGI SAS\Documents\GitHub\SGI\outputMDVR\general_information_report_2024_08_01_00_00_00_2024_08_02_00_00_00_1722540147.xlsx'
# file_mdvr2 = r'C:\Users\SGI SAS\Documents\GitHub\SGI\outputMDVR\overspeeds_report_2024_08_01_00_00_00_2024_08_02_00_00_00_1722540155.xlsx'
# file_ubicom1 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputUbicom\ReporteDiario.xls"
# file_ubicom2 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputUbicom\Estacionados.xls"
# file_securitrac = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputSecuritrac\exported-excel.xls"
# file_wialon1 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputWialon\JTV645_INFORME_GENERAL_TM_V1.0_2024-08-01_14-24-44.xlsx"
# file_wialon2 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputWialon\LPN816_INFORME_GENERAL_TM_V1.0_2024-08-01_14-24-53.xlsx"  
# file_wialon3 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputWialon\LPN821_INFORME_GENERAL_TM_V1.0_2024-08-01_14-25-02.xlsx"
# file_ituran1 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputIturan\report.csv"
# file_ituran2 = r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputIturan\report (1).csv"
# # datos = functionalidad.extraerMDVR(file1, file2)
# # print(datos)

# # rpa.ejecutarRPAMDVR()
# # rpa.ejecutarRPAIturan()
# # rpa.ejecutarRPASecuritrac()
# # rpa.ejecutarRPAUbicar()
# # rpa.ejecutarRPAWialon()
# # rpa.ejecutarRPAUbicom()


# # datos = functionalidad.extraerIturan(file_ituran1, file_ituran2)
# # print(datos)


# # def OdomIturan(file):
# #     # Leer el archivo de Excel
# #     od = pd.read_csv(file)

# #     # Extraer la placa y el odómetro

# #     df = od[['V_PLATE_NUMBER', 'END_ODOMETER']]

# #     # Renombrar las columnas

# #     df.columns = ['PLACA', 'KILOMETRAJE']

# #     # Crear el diccionario con el formato requerido

# #     datos = df.to_dict('records')
# #     return datos

# # odom = OdomIturan(r'C:\Users\SGI SAS\Documents\GitHub\SGI\outputIturan\report (3).csv')

# # print(odom)

# # functionalidad.extraerIturan(file_ituran1, file_ituran2)

# # wialon.rpaWialon()

# # x, y = consulta.tablaCorreoLaboral()

# # print(x)


# rutasLaboral = {'securitrac': r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputSecuritrac\exported-excel.xls",

#             'mdvr': r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputMDVR\drives_stops_drivers_report_2024_08_09_00_00_00_2024_08_10_00_00_00_1723225380.xlsx",

#             'ituran': r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputIturan\report (2).csv",

#             'ubicar': r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputUbicar\drives_stops_report_2024_08_09_00_00_00_2024_08_10_00_00_00_1723225435.xlsx",

#             'wialon': [r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputWialon\JTV645_INFORME_GENERAL_TM_V1.0_2024-08-09_12-45-06.xlsx", r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputWialon\LPN816_INFORME_GENERAL_TM_V1.0_2024-08-09_12-45-14.xlsx", r"C:\Users\SGI SAS\Documents\GitHub\SGI\outputPrueba\outputWialon\LPN821_INFORME_GENERAL_TM_V1.0_2024-08-09_12-45-23.xlsx"]

#             }

# df = functionalidad.fueraLaboralTodos(rutasLaboral)

# print(df)









