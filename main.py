import sys
from forms.ituranForm import rpaIturan
from forms.MDVRForm import rpaMDVR
from forms.securitracForm import rpaSecuritrac
from forms.ubicarForm import rpaUbicar
from forms.ubicomForm import rpaUbicom
from forms.wialonForm import rpaWialon
from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput, enviarCorreoConductor, enviarCorreoPlataforma
from persistence.archivoExcel import crear_excel, actualizarInfractores, odomUbicar, OdomIturan, actualizarOdom
from persistence.scriptMySQL import actualizarKilometraje, actualizarSeguimientoSQL 
##### FALTA ACTUALIZAR INFRACTORES SQL Y SEGUIMIENTO SQL
##### FALTA CORREO CON RPA NO FUNCIONÓ.

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
        rpaMDVR()
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
        rpaUbicar()
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
        rpaWialon()
    except:
        print("Hubo un error en el acceso por el internet.")
        enviarCorreoPlataforma("Wialon")


    ####################################
    ####### Creación de informes #######
    ####################################


    # Actualización de seguimiento
    crear_excel()

    # Actualización de infractores
    actualizarInfractores

    # Actualización del odómetro
    odomUbicar()
    OdomIturan()
    actualizarOdom()

    # Actualización de indicadores
    ################################################

    # Conexión con la base de datos
    actualizarKilometraje()
    actualizarSeguimientoSQL()


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



