import sys, os
from datetime import datetime
import pandas as pd
from db.consultasImportantes import ConsultaImportante
from forms.rpaCompleto import RPA
from persistence.extraerExcel import Extracciones
from persistence.insertarSQL import FuncionalidadSQL
from persistence.archivoExcel import FuncionalidadExcel
from util.correosVehiculares import CorreosVehiculares
from util.tratadoArchivos import TratadorArchivos

def main(self):
    """
    Ejecuta todos los códigos de la RPA en orden.
    """


    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    # Ituran
    archivoIturan1, archivoIturan2, archivoIturan3, archivoIturan4 = RPA().ejecutarRPAIturan()

    # MDVR
    archivoMDVR1,archivoMDVR2, archivoMDVR3 = RPA().ejecutarRPAMDVR()

    # Securitrac
    archivoSecuritrac = RPA().ejecutarRPASecuritrac()

    # Ubicar
    archivoUbicar1,archivoUbicar2,archivoUbicar3 = RPA().ejecutarRPAUbicar()

    # Ubicom
    archivoUbicom1, archivoUbicom2 = RPA().ejecutarRPAUbicom()

    # Wialon
    archivoWialon1, archivoWialon2, archivoWialon3 = RPA().ejecutarRPAWialon()

    # if all(ele == "Ejecutado" for ele in tablaEstadosTotales['estado'].values) == True:
    #     pass
    # else:
    #     sys.exit() # En caso de que no todos sean Ejecutado, no se sigue.

    ####################################
    ##### Verificar estados Finales ####
    ####################################

# ############# 11:45pm
#     tablaEstadosTotales = ConsultaImportante().verificarEstadosFinales()
#     tablaEstadosTotales = pd.DataFrame(tablaEstadosTotales, columns=['plataforma', 'estado']).set_index('plataforma')
#     for plataforma in tablaEstadosTotales.index:
#         estado = tablaEstadosTotales.loc[plataforma]['estado']
#         if estado == "Ejecutado":
#             pass
#         else:
#             ConsultaImportante().registrarError(plataforma)
#             CorreosVehiculares().enviarCorreoPlataforma(plataforma)
#             pass
    

#     ########### no
#     if all(ele == "Ejecutado" for ele in tablaEstadosTotales['estado'].values) == True:
#         pass
#     else:
#         sys.exit() # En caso de que no todos sean Ejecutado, no se sigue.


#     # si todo fue ejecutado ya y la hora es menor a las 11:45pm, sys exit
#     # si no todo fue ejecutado y la hora es menor a las 11:45pm, sys exit
#     # si no todo fue ejecutado y la hora es mayor o igual a las 11:45pm
#     # sigue todo.
#     # si todo fue ejecutado ya y la hora es mayor o igual a las 11:45pm
#     # se puede actualizar la tabla de estados pero no se hace nada más



    ####################################
    ####### Creación de informes #######
    ####################################


    archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

    # Actualización de seguimiento
    df_exist = Extracciones().crear_excel(archivoMDVR1,archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)

    # Actualización de infractores
    Extracciones().actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

    # Actualización del odómetro
    Extracciones().actualizarOdom(archivoSeguimiento, archivoIturan3, archivoUbicar1)

    # Actualización de indicadores
    df_diario = Extracciones().dfDiario(df_exist)
    Extracciones().actualizarIndicadoresTotales(df_diario, archivoSeguimiento)
    Extracciones().actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

    # Conexión con la base de datos
    FuncionalidadSQL().actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, archivoMDVR1, archivoMDVR2, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3)
    FuncionalidadSQL().actualizarInfractoresSQL(archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)
    

    ####################################
    ##### Fuera de horario laboral #####
    ####################################

    rutasLaboral = {'securitrac': archivoSecuritrac,

                'mdvr': archivoMDVR2,

                'ituran': archivoIturan4,

                'ubicar': archivoUbicar2,

                'wialon': [archivoWialon1, archivoWialon2, archivoWialon3]

                }
    
    # Excel
    fueraHorarioLaboral = FuncionalidadExcel().fueraLaboralTodos(rutasLaboral)
    Extracciones().actualizarFueraLaboral(archivoSeguimiento, fueraHorarioLaboral)

    # SQL
    FuncionalidadSQL().sqlFueraLaboral(fueraHorarioLaboral)

    ####################################
    ######### Envío de correos #########
    ####################################


    # Enviar correo al personal de SGI.
    CorreosVehiculares().enviarCorreoPersonal()

    # Enviar correo específico a los conductores con excesos de velocidad.
    CorreosVehiculares().enviarCorreoConductor()
    
    # Enviar correo al personal de SGI de vehículos fuera de horario laboral.
    CorreosVehiculares().enviarCorreoLaboral()

    ####################################
    ######### Borrado y salida #########
    ####################################


    # Eliminar las carpetas del output ya que se tiene toda la información.
    # print("Eliminando archivos")
    # time.sleep(10)
    # TratadorArchivos().eliminarArchivosOutput()

    # Actualización de la tabla de estados.
    hora = int(str(datetime.now().hour) + str(datetime.now().minute))

    if hora >= 2345:
        ConsultaImportante().actualizarTablaEstados()
    else:
        pass

    # Salida del sistema.
    sys.exit()


if __name__=='__main__':
    main()



