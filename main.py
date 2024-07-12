import sys
from forms.ituranForm import IturanDatos
# from forms.MDVRForm import rpaMDVR
# from forms.securitracForm import rpaSecuritrac
# from forms.ubicarForm import rpaUbicar
# from forms.ubicomForm import rpaUbicom
from forms.wialonForm import WialonDatos
from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput


def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """


    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    # Ituran
    try:
        IturanDatos.rpaIturan()
    except:
        eliminarArchivosOutput()
        print("Hubo un error en el acceso por el internet.")
        sys.exit()

    # # MDVR
    # try:
    #     MDVRDatos.rpaMDVR()
    # except:
    #     eliminarArchivosOutput()
    #     print("Hubo un error en el acceso por el internet.")
    #     sys.exit()
    
    # # Securitrac
    # try:
    #     securitracDatos.rpaSecuritrac()
    # except:
    #     eliminarArchivosOutput()
    #     print("Hubo un error en el acceso por el internet.")
    #     sys.exit()

    # # Ubicar
    # try:
    #     ubicarDatos.rpaUbicar()
    # except:
    #     eliminarArchivosOutput()
    #     print("Hubo un error en el acceso por el internet.")
    #     sys.exit()

    # # Ubicom
    # try:
    #     ubicomDatos.rpaUbicom()
    # except:
    #     eliminarArchivosOutput()
    #     print("Hubo un error en el acceso por el internet.")
    #     sys.exit()

    # # Wialon
    # try:
    #     wialonDatos.rpaWialon()
    # except:
    #     eliminarArchivosOutput()
    #     print("Hubo un error en el acceso por el internet.")
    #     sys.exit()


    ####################################
    ######## Realizar informes #########
    ####################################


    # crear excel
    # actualizar infractores
    # actualizar odometro


    # # Enviar correo.
    #enviarCorreoPersonal()
    
    # Eliminar las carpetas del output ya que se tiene toda la información.
    #eliminarArchivosOutput()

    # Salida del sistema.
    sys.exit()

if __name__=='__main__':
    main()



