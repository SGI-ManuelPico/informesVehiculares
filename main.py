import sys, os, time
from forms.rpaCompleto import RPA
from persistence.extraerExcel import Extracciones
from persistence.insertarSQL import FuncionalidadSQL
from util.correosVehiculares import CorreosVehiculares
from util.tratadoArchivos import TratadorArchivos

def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """


    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    # Ituran
    try:
        archivoIturan1, archivoIturan2, archivoIturan3 = RPA().ejecutarRPAIturan()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Ituran.")
        #enviarCorreoPlataforma("Ituran")

    # MDVR
    try:
        archivoMDVR1,archivoMDVR2, archivoMDVR3 = RPA().ejecutarRPAMDVR()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a MDVR.")
        #enviarCorreoPlataforma("MDVR")
    
    # Securitrac
    try:
        archivoSecuritrac = RPA().ejecutarRPASecuritrac()
    except:
        print("Hubo un error en el acceso por el internetpara ingresar a Securitrac.")
        archivoSecuritrac = os.getcwd() + "\\outputSecuritrac\\exported-excel.xls"
        #enviarCorreoPlataforma("Securitrac")

    # Ubicar
    try:
        archivoUbicar1,archivoUbicar2,archivoUbicar3 = RPA().ejecutarRPAUbicar()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Ubicar.")
        #enviarCorreoPlataforma("Ubicar")

    # Ubicom
    try:
        archivoUbicom1, archivoUbicom2 = RPA().ejecutarRPAUbicom()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Ubicom.")
        #enviarCorreoPlataforma("Ubicom")

    # Wialon
    try:
        archivoWialon1, archivoWialon2, archivoWialon3 = RPA().ejecutarRPAWialon()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Wialon.")
        #enviarCorreoPlataforma("Wialon")


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
    ######### Envío de correos #########
    ####################################


    # Enviar correo al personal de SGI.
    CorreosVehiculares().enviarCorreoPersonal()

    # Enviar correo específico a los conductores con excesos de velocidad.
    CorreosVehiculares().enviarCorreoConductor()
    

    ####################################
    ######### Borrado y salida #########
    ####################################


    # Eliminar las carpetas del output ya que se tiene toda la información.
    # print("Eliminando archivos")
    # time.sleep(10)
    # TratadorArchivos().eliminarArchivosOutput()
    print("Gatitos24")

    # Salida del sistema.
    sys.exit()

if __name__=='__main__':
    main()



