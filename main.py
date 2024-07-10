import sys
from forms.ituranForm import rpaIturan
from forms.MDVRForm import rpaMDVR
from forms.securitracForm import rpaSecuritrac
from forms.ubicarForm import rpaUbicar
from forms.ubicomForm import rpaUbicom
from forms.wialonForm import rpaWialon
from util.funcionalidadVehicular import enviarCorreo, eliminarArchivosOutput


def main():

    # Realizar los RPA en orden.
    rpaIturan()
    rpaMDVR()
    rpaSecuritrac()
    rpaUbicar()
    rpaUbicom()
    rpaWialon()

    # Ejecutar informe.

    # Enviar correo.
    enviarCorreo()
    
    # Eliminar las carpetas del output ya que se tiene toda la información.
    eliminarArchivosOutput()

    # Salida del sistema.
    sys.exit()

if __name__=='__main__':
    main()



