import sys
#from forms.ituranForm import IturanDatos
# from forms.MDVRForm import rpaMDVR
# from forms.securitracForm import rpaSecuritrac
# from forms.ubicarForm import rpaUbicar
# from forms.ubicomForm import rpaUbicom
#from forms.wialonForm import WialonDatos
#from util.funcionalidadVehicular import enviarCorreoPersonal, eliminarArchivosOutput
from persistence.updateSQL import ActualizarBD


def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """
    #PRUEBA GABRIEL
    actualizar = ActualizarBD()

    #Actualizar Seguimiento
    file_ituran1 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ituran1.csv"
    file_ituran2 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ituran2.csv"
    file_MDVR1   = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_MDVR1.xls"
    file_MDVR2   = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_MDVR2.xlsx"
    file_Ubicar1 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ubicar1.xlsx"
    file_Ubicar2 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ubicar2.xlsx"
    file_Ubicom1 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ubicom1.xls"
    file_Ubicom2 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ubicom2.xls"
    file_Securitrac = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_securitrac.xls"
    file_Wialon1 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_Wialon1.xlsx"
    file_Wialon2 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_Wialon2.xlsx"
    file_Wialon3 = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_Wialon3.xlsx"
    #Actualizar Kilometraje
    file_ituranKM = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_ituranKM.csv"
    file_UbicarKM = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\file_UbicarKM.xlsx"
    #Infractores
    file_IturanInfractores  = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\ituranInfracciones.xls"
    file_MDVRInfractores    = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\MDVRinfractores.xlsx"
    file_UbicarInfractores  = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\ubicarInfractores.xlsx"
    file_Wialon1Infractores = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\wialonInfrac1Hist.xlsx"
    file_Wialon2Infractores = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\wialonInfrac2Hist.xlsx"
    file_Wialon3Infractores = r"C:\Users\Soporte\Documents\GitHub\SGI\pruebaVehiculosExcel\wialonInfrac3Hist.xlsx"
    file_SecuritracInfractores  = r"pruebaVehiculosExcel\securitracInfractores.xls"


    #actualizar.actualizarSeguimientoSQL(file_ituran1, file_ituran2, file_MDVR1, file_MDVR2, file_Ubicar1, file_Ubicar2, file_Ubicom1, file_Ubicom2, file_Securitrac, file_Wialon1, file_Wialon2, file_Wialon3)
    #actualizar.actualizarKilometraje(file_ituranKM, file_UbicarKM)
    #actualizar.actualizarInfractoresSQL(file_IturanInfractores, file_MDVRInfractores, file_UbicarInfractores, file_Wialon1Infractores, file_Wialon2Infractores, file_Wialon3Infractores, file_SecuritracInfractores)

    ####################################
    ###### RPA por cada plataforma #####
    ####################################


    #Ituran
    # try:
    #     IturanDatos.rpaIturan()
    # except:
    #     eliminarArchivosOutput()
    #     print("Hubo un error en el acceso por el internet.")
    #     sys.exit()

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


    #Data Frame
    #Crear Excel
    # actualizar infractores
    # actualizar odometro


    # # Enviar correo.
    #enviarCorreoPersonal()
    
    # Eliminar las carpetas del output ya que se tiene toda la información.
    #eliminarArchivosOutput()

    # Salida del sistema.
    #sys.exit()

if __name__=='__main__':
    main()



