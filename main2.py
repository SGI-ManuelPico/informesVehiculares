import sys, os, time
from forms.ituranForm import DatosIturan
from forms.MDVRForm import DatosMDVR
from forms.securitracForm import DatosSecuritrac
from forms.ubicarForm import DatosUbicar
from forms.ubicomForm import DatosUbicom
from forms.wialonForm import DatosWialon
from util.correosVehiculares import CorreosVehiculares
from persistence.archivoExcel import FuncionalidadExcel
from persistence.extraerExcel import Extracciones
from persistence.scriptMySQL import actualizarKilometraje, actualizarSeguimientoSQL, actualizarInfractoresSQL
from persistence.estadoPlataforma import EstadoPlataforma
from util.tratadoArchivos import TratadorArchivos

### Como estas funciones solo se van a ejecutar en el main, las defino acá. ###

estado_plataforma = EstadoPlataforma()


def ejecutarTareaRPA(plataforma, funcionRPA):
    try:
        estado_plataforma.actualizar_estado(plataforma, 'En ejecucion')
        resultado = funcionRPA()
        estado_plataforma.actualizar_estado(plataforma, 'Finalizado')
        return resultado
    except Exception as e:
        print(f"Hubo un error en el acceso por el internet para ingresar a {plataforma}.")
        estado_plataforma.actualizar_estado(plataforma, 'Error')
        CorreosVehiculares.enviarCorreoPlataforma(plataforma)

# Verifica si hay alguna plataforma que tenga estado 'error' en la tabla estadoPlataforma. Si los hay, retorna True, si no, retorna False.
def checkErrores():
    status = estado_plataforma.verificar_estado()
    for plataforma, estado in status:
        if estado == 'Error':
            return True
    return False

# Si hay una plataforma con estado 'error' en la tabla estadoPlataforma, ejecuta el RPA para esa plataforma específica.
def retryErrores(plataformas, resultados):
    status = estado_plataforma.verificar_estado()
    for plataforma, estado in status:
        if estado == 'Error':
            resultados[plataforma] = ejecutarTareaRPA(plataforma, dict(plataformas)[plataforma])


def main():
    
    """
    Ejecuta todos los códigos de la RPA en orden. Guardamos los resultados de las RPA en un diccionario que tiene como llaves
    los nombres de las plataformas y como valores las rutas que se generan con cada RPA.
    """
     
    resultados = {}

    # Creamos una lista que contiene tuplas con el nombre de la plataforma y su funcion RPA correspondiente.
    plataformas = [
        ('Ituran', DatosIturan.rpaIturan),
        ('MDVR',  DatosMDVR.rpaMDVR),
        ('Securitrac', DatosSecuritrac.rpaSecuritrac),
        ('Ubicar', DatosUbicar.rpaUbicar),
        ('Ubicom', DatosUbicom.rpaUbicom),
        ('Wialon', DatosWialon.rpaWialon)
    ]

    # Ejectuamos el RPA para cada plataforma y guardamos los resultados en el diccionario.
    
    for plataforma, funcionRPA in plataformas:
        resultados[plataforma] = ejecutarTareaRPA(plataforma, funcionRPA)
    
    # Primer reintento después de 15 minutos solo si hay al menos una plataforma con estado 'Error'
    if checkErrores():
        time.sleep(15 * 60)
        retryErrores(plataformas, resultados)
        
        # Segundo reintento después de 15 minutos solo si hay al menos una plataforma con estado 'Error'
        if checkErrores():
            time.sleep(15 * 60)
            retryErrores(plataformas, resultados)

    # Verificar estados finales y registrar errores en la tabla 'error'
    estadoFinal = estado_plataforma.verificar_estado()
    existe_error = False
    for plataforma, estado in estadoFinal:
        if estado == 'Error':
            estado_plataforma.log_error(plataforma)
            existe_error = True

    # Si no hay errores en las plataformas, reseteamos el estado de todas las plataformas en la tabla 'estadoPlataforma'
    if not existe_error:
        estado_plataforma.reset_estados()

    # Rutas.
    archivoIturan1, archivoIturan2, archivoIturan3 = resultados.get('Ituran', (None, None, None))
    archivoMDVR1, archivoMDVR2, archivoMDVR3 = resultados.get('MDVR', (None, None, None))
    archivoSecuritrac = resultados.get('Securitrac', None)
    archivoUbicar1, archivoUbicar2, archivoUbicar3 = resultados.get('Ubicar', (None, None, None))
    archivoUbicom1, archivoUbicom2 = resultados.get('Ubicom', (None, None))
    archivoWialon1, archivoWialon2, archivoWialon3 = resultados.get('Wialon', (None, None, None))

    ####################################
    ###### Creación de informes #####
    ####################################

    archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

    try:
        df_exist = Extracciones.crear_excel(archivoMDVR1,archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)

        Extracciones.actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

        Extracciones.actualizarOdom(archivoSeguimiento, archivoIturan3, archivoUbicar1)

        df_diario = Extracciones.dfDiario(df_exist)

        Extracciones.actualizarIndicadoresTotales(df_diario, archivoSeguimiento)

        Extracciones.actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

    except Exception as e:
        print(f"Error al procesar los archivos de Excel: {e}")
    
    ####################################
    ###### Actualización de MySQL ######
    ####################################

    try:
        actualizarKilometraje()
        actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, archivoMDVR1, archivoMDVR2, archivoMDVR3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)
        actualizarInfractoresSQL(archivoIturan2, archivoMDVR3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)
    except Exception as e:
        print(f"Error al actualizar la base de datos MySQL: {e}")

    ####################################
    ########## Envío de correos ########
    ####################################

    try:
        CorreosVehiculares.enviarCorreoPersonal()
        CorreosVehiculares.enviarCorreoConductor()
    except Exception as e:
        print(f"Error al enviar correos: {e}")

        # Eliminar las carpetas del output ya que se tiene toda la información.
    print("Eliminando archivos")
    time.sleep(10)
    TratadorArchivos.eliminarArchivosOutput()

    # Salida del sistema.
    sys.exit()

if __name__ == "__main__":
    main()
