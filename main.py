import sys
from forms.ituranForm import rpaIturan
from forms.MDVRForm import rpaMDVR
from forms.securitracForm import rpaSecuritrac
from forms.ubicarForm import rpaUbicar
from forms.ubicomForm import rpaUbicom
from forms.wialonForm import rpaWialon
from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput

def main():

    # Realizar los RPA en orden.
    # rpaIturan()
    # rpaMDVR()
    # rpaSecuritrac()
    # rpaUbicar()
    # rpaUbicom()
    # rpaWialon()

    # # SE DEBERÍA COLOCAR COMO TRY EXCEPT POR SI ACASO EL INTERNET LO DAÑA y que solo ejecute si no hay archivos ahí.
    # # TAMBIÉN ALGO PARA QUE MIRE SI EL INTERNET ESTÁ BIEN PARA CORRERLO, AUNQUE CREO QUE ESTO YA SERÍA DENTRO DEL BAT.

    # # Ejecutar informe.


    # # Enviar correo.
    enviarCorreoPersonal()
    
    # Eliminar las carpetas del output ya que se tiene toda la información.
    #eliminarArchivosOutput()

    # Salida del sistema.
    sys.exit()

if __name__=='__main__':
    main()



