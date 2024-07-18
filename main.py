import sys, os
from forms.ituranForm import rpaIturan, archivoIturan1, archivoIturan2, archivoIturan3
from forms.MDVRForm import MDVRDatos
from forms.securitracForm import rpaSecuritrac, archivoSecuritrac
from forms.ubicarForm import ubicarDatos
from forms.ubicomForm import rpaUbicom, archivoUbicom1, archivoUbicom2
from forms.wialonForm import wialonDatos
from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput, enviarCorreoConductor, enviarCorreoPlataforma
from persistence.archivoExcel import crear_excel, actualizarInfractores, actualizarOdom, actualizarIndicadoresTotales, actualizarIndicadores, dfDiario
from persistence.scriptMySQL import actualizarKilometraje, actualizarSeguimientoSQL 
##### FALTA ACTUALIZAR INFRACTORES SQL Y SEGUIMIENTO SQL
##### FALTA ACTUALIZAR INDICADORES Y TOTAL


def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """


    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    # Ituran
    try:
        rpaIturan()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("Ituran")

    # MDVR
    try:
        MDVRDatos.rpaMDVR()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("MDVR")
    
    # Securitrac
    try:
        rpaSecuritrac()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("Securitrac")

    # Ubicar
    try:
        ubicarDatos.rpaUbicar()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("Ubicar")

    # Ubicom
    try:
        rpaUbicom()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("Ubicom")

    # Wialon
    try:
        wialonDatos.rpaWialon()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("Wialon")


    ####################################
    ####### Creación de informes #######
    ####################################


    archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

    # Actualización de seguimiento
    df_exist = crear_excel(MDVRDatos.archivoMDVR1,MDVRDatos.archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, wialonDatos.archivoWialon1, wialonDatos.archivoWialon2, wialonDatos.archivoWialon3, ubicarDatos.archivoUbicar1, ubicarDatos.archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)
    df_diario = dfDiario(df_exist)

    # Actualización de infractores
    actualizarInfractores(archivoSeguimiento, archivoIturan2, MDVRDatos.archivoMDVR3, ubicarDatos.archivoUbicar3, wialonDatos.archivoWialon1, wialonDatos.archivoWialon2, wialonDatos.archivoWialon3, archivoSecuritrac)

    # Actualización del odómetro
    actualizarOdom(archivoSeguimiento, archivoIturan3, ubicarDatos.archivoUbicar1)

    # Actualización de indicadores
    actualizarIndicadoresTotales(df_diario, archivoSeguimiento)
    actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

    # Conexión con la base de datos
    actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, MDVRDatos.archivoMDVR1, MDVRDatos.archivoMDVR2, ubicarDatos.archivoUbicar1, ubicarDatos.archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSecuritrac, wialonDatos.archivoWialon1, wialonDatos.archivoWialon2, wialonDatos.archivoWialon3)


    ####################################
    ######### Envío de correos #########
    ####################################


    # Enviar correo al personal de SGI.
    enviarCorreoPersonal()

    # Enviar correo específico a los conductores con excesos de velocidad.
    enviarCorreoConductor()
    

    ####################################
    ######### Borrado y salida #########
    ####################################


    # Eliminar las carpetas del output ya que se tiene toda la información.
    #eliminarArchivosOutput()

    # Salida del sistema.
    sys.exit()

if __name__=='__main__':
    main()



