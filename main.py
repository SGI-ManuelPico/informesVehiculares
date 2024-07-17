import sys, os
from forms.ituranForm import rpaIturan, archivoIturan1, archivoIturan2, archivoIturan3
from forms.MDVRForm import MDVRDatos
from forms.securitracForm import rpaSecuritrac, archivoSecuritrac
from forms.ubicarForm import rpaUbicar
from forms.ubicomForm import rpaUbicom, archivoUbicom1, archivoUbicom2
from forms.wialonForm import rpaWialon
from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput, enviarCorreoConductor, enviarCorreoPlataforma
from persistence.archivoExcel import crear_excel, actualizarInfractores, actualizarOdom
from persistence.scriptMySQL import actualizarKilometraje, actualizarSeguimientoSQL 
##### FALTA ACTUALIZAR INFRACTORES SQL Y SEGUIMIENTO SQL
##### FALTA ACTUALIZAR INDICADORES Y TOTAL

print(MDVRDatos.archivoMDVR1)
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


    archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

    # Actualización de seguimiento
    crear_excel(MDVRDatos.archivoMDVR1,MDVRDatos.archivoMDVR3, archivoIturan3, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)

    # Actualización de infractores
    actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

    # Actualización del odómetro
    actualizarOdom(archivoSeguimiento, archivoIturan3, archivoUbicar1)

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



