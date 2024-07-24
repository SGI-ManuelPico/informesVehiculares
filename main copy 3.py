import sys, os, time
from datetime import datetime
import pandas as pd
from db.consultasImportantes import ConsultaImportante
from forms.rpaCompleto import RPA
from persistence.extraerExcel import Extracciones
from persistence.insertarSQL import FuncionalidadSQL
from persistence.archivoExcel import FuncionalidadExcel
from util.correosVehiculares import CorreosVehiculares
from util.tratadoArchivos import TratadorArchivos

def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """

    ########################
    ### Consulta inicial ###
    ########################

    tablaEstadosTotales = ConsultaImportante().verificarEstadosFinales()
    tablaEstadosTotales = pd.DataFrame(tablaEstadosTotales, columns=['plataforma', 'estado']).set_index('plataforma')


    ######## Todo bien inicio
    if all(ele == "Ejecutado" for ele in tablaEstadosTotales['estado'].values) == True:
        sys.exit()
    else:
        pass # En caso de que no todos sean Ejecutado, no se sigue.


    ### En caso de que se siga, revisar qué plataformas tuvieron errores.
    plataformasFallidas = []
    for plataforma in tablaEstadosTotales.index: #Verifica qué plataformas definitivamente tuvieron errores.
        estado = tablaEstadosTotales.loc[plataforma]['estado']
        if estado == "Ejecutado":
            pass
        else:
            plataformasFallidas.append(plataforma)


    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    for plataforma in plataformasFallidas: # Realiza los RPA de las plataformas fallidas
        if plataforma in plataformasFallidas and plataforma == "Ituran":
            archivoIturan1, archivoIturan2, archivoIturan3, archivoIturan4 = RPA().ejecutarRPAIturan()
        else:
            pass

        if plataforma in plataformasFallidas and plataforma == "Securitrac":
            archivoSecuritrac = RPA().ejecutarRPASecuritrac()
        else:
            pass

        if plataforma in plataformasFallidas and plataforma == "MDVR":
            archivoMDVR1,archivoMDVR2, archivoMDVR3 = RPA().ejecutarRPAMDVR()
        else:
            pass

        if plataforma in plataformasFallidas and plataforma == "Ubicar":
            archivoUbicar1,archivoUbicar2,archivoUbicar3 = RPA().ejecutarRPAUbicar()
        else:
            pass

        if plataforma in plataformasFallidas and plataforma == "Ubicom":
            archivoUbicom1, archivoUbicom2 = RPA().ejecutarRPAUbicom()
        else:
            pass

        if plataforma in plataformasFallidas and plataforma == "Wialon":
            archivoWialon1, archivoWialon2, archivoWialon3 = RPA().ejecutarRPAWialon()
        else:
            pass


    ####################################
    ### Verificar estados intermedios ##
    ####################################


    # Si alguna plataforma falló, pero se arregló, se ejecutará el archivo. Si no pasa esto, se sale.
    if all(ele == "Ejecutado" for ele in tablaEstadosTotales['estado'].values) == True:
        pass
    else:
        sys.exit() # En caso de que no todos sean Ejecutado, no se sigue.

    
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
    ######## Salida del sistema ########
    ####################################


    sys.exit()


if __name__=='__main__':
    main()



