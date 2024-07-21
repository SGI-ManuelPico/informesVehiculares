import sys, os
from forms.ituranForm import rpaIturan
from forms.MDVRForm import rpaMDVR
from forms.securitracForm import rpaSecuritrac
from forms.ubicarForm import rpaUbicar
from forms.ubicomForm import rpaUbicom
from forms.wialonForm import rpaWialon
from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput, enviarCorreoConductor, enviarCorreoPlataforma
from persistence.archivoExcel import crear_excel, actualizarInfractores, actualizarOdom, actualizarIndicadoresTotales, actualizarIndicadores, dfDiario
from persistence.scriptMySQL import actualizarKilometraje, actualizarSeguimientoSQL, actualizarInfractoresSQL

def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """


    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    # Ituran
    try:
        archivoIturan1, archivoIturan2, archivoIturan3 = rpaIturan()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Ituran.")
        #enviarCorreoPlataforma("Ituran")

    # MDVR
    try:
        archivoMDVR1,archivoMDVR2, archivoMDVR3 = rpaMDVR()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a MDVR.")
        #enviarCorreoPlataforma("MDVR")
    
    # Securitrac
    try:
        archivoSecuritrac = rpaSecuritrac()
    except:
        print("Hubo un error en el acceso por el internetpara ingresar a Securitrac.")
        archivoSecuritrac = os.getcwd() + "\\outputSecuritrac\\exported-excel.xls"
        #enviarCorreoPlataforma("Securitrac")

    # Ubicar
    try:
        archivoUbicar1,archivoUbicar2,archivoUbicar3 = rpaUbicar()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Ubicar.")
        #enviarCorreoPlataforma("Ubicar")

    # Ubicom
    try:
        archivoUbicom1, archivoUbicom2 = rpaUbicom()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Ubicom.")
        #enviarCorreoPlataforma("Ubicom")

    # Wialon
    try:
        archivoWialon1, archivoWialon2, archivoWialon3 = rpaWialon()
    except:
        print("Hubo un error en el acceso por el internet para ingresar a Wialon.")
        #enviarCorreoPlataforma("Wialon")


    ####################################
    ####### Creación de informes #######
    ####################################


    # archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

    # # Actualización de seguimiento
    # crear_excel(mdvr_file1= archivoMDVR1,mdvr_file2=archivoMDVR3, ituran_file=archivoIturan1, ituran_file2=archivoIturan2, securitrac_file=archivoSecuritrac, wialon_file1=archivoWialon1, wialon_file2=archivoWialon2, wialon_file3=archivoWialon3, ubicar_file1=archivoUbicar1, ubicar_file2=archivoUbicar2, ubicom_file1=archivoUbicom1, ubicom_file2=archivoUbicom2, output_file=archivoSeguimiento)

    # # Actualización de infractores
    # actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

    # # Actualización del odómetro
    # actualizarOdom(archivoSeguimiento, archivoIturan3, archivoUbicar1)

    # # Actualización de indicadores
    # df_exist = crear_excel(archivoMDVR1,archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)
    # df_diario = dfDiario(df_exist)
    # actualizarIndicadoresTotales(df_diario, archivoSeguimiento)
    # actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

    # Conexión con la base de datos
    actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, archivoMDVR1, archivoMDVR2, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3)
    actualizarInfractoresSQL(archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)


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
    eliminarArchivosOutput()

    # Salida del sistema.
    sys.exit()

if __name__=='__main__':
    main()



