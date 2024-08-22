import sys, os, time
import logging
from datetime import datetime
import pandas as pd
from db.consultasImportantes import ConsultaImportante
from forms.rpaCompleto import RPA
from persistence.extraerExcel import Extracciones
from persistence.insertarSQL import FuncionalidadSQL
from persistence.archivoExcel import FuncionalidadExcel
from util.correosVehiculares import CorreosVehiculares
from util.tratadoArchivos import TratadorArchivos

def main():
    """
    Ejecuta todos los códigos de la RPA en orden.
    """
    try:    

        # Actualización de la tabla de estados.
        hora = int(str(datetime.now().hour) + str(datetime.now().minute))
        print(hora)

        if hora <= 2310:

            TratadorArchivos().eliminarArchivosOutput()
            ConsultaImportante().actualizarTablaEstados()  

        
            ########################
            ### Consulta inicial ###
            ########################


            tablaEstadosTotales = ConsultaImportante().verificarEstadosFinales()
            tablaEstadosTotales = pd.DataFrame(tablaEstadosTotales, columns=['plataforma', 'estado']).set_index('plataforma')

            listaEstadosTotales = []


            ####################################
            ###### RPA por cada plataforma #####
            ####################################


            # Ituran
            archivoIturan1, archivoIturan2, archivoIturan3, archivoIturan4 = RPA().ejecutarRPAIturan()
            if archivoIturan4 == os.getcwd() + r"\archivoFicticio.csv":
                listaEstadosTotales.append('Error')
            else:
                listaEstadosTotales.append('Ejecutado')

            # MDVR
            archivoMDVR1,archivoMDVR2, archivoMDVR3 = RPA().ejecutarRPAMDVR()
            if archivoMDVR1 == os.getcwd() + r"\archivoFicticio.xlsx":
                listaEstadosTotales.append('Error')
            else:
                listaEstadosTotales.append('Ejecutado')

            # Securitrac
            archivoSecuritrac = RPA().ejecutarRPASecuritrac()
            if archivoSecuritrac == os.getcwd() + r"\archivoFicticio.xls":
                listaEstadosTotales.append('Error')
            else:
                listaEstadosTotales.append('Ejecutado')

            # Ubicar
            archivoUbicar1,archivoUbicar2,archivoUbicar3 = RPA().ejecutarRPAUbicar()
            if archivoUbicar1 == os.getcwd() + r"\archivoFicticio.xlsx":
                listaEstadosTotales.append('Error')
            else:
                listaEstadosTotales.append('Ejecutado')

            # Ubicom
            archivoUbicom1, archivoUbicom2 = RPA().ejecutarRPAUbicom()
            if archivoUbicom1 == os.getcwd() + r"\archivoFicticio.xls":
                listaEstadosTotales.append('Error')
            else:
                listaEstadosTotales.append('Ejecutado')

            # Wialon
            archivoWialon1, archivoWialon2, archivoWialon3 = RPA().ejecutarRPAWialon()
            if archivoWialon1 == os.getcwd() + r"\archivoFicticio.xlsx":
                listaEstadosTotales.append('Error')
            else:
                listaEstadosTotales.append('Ejecutado')


            ####################################
            #### Verificar estados iniciales ###
            ####################################


            # Si alguna plataforma falló, pero se arregló, se ejecutará el archivo. Si no pasa esto, se sale.

            if all(ele == "Ejecutado" for ele in listaEstadosTotales) == False:
                sys.exit() # En caso de que no todos sean Ejecutado, no se sigue.

            print("sigue")


            ####################################
            ####### Creación de informes #######
            ####################################


            archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

            # Actualización de seguimiento
            df_exist = Extracciones().crear_excel(archivoMDVR1,archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)

            # Actualización de infractores
            Extracciones().actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

            # Actualización del odómetro
            Extracciones().actualizarOdom(archivoSeguimiento, archivoIturan4, archivoUbicar1)

            # Actualización de indicadores
            df_diario = Extracciones().dfDiario(df_exist)
            Extracciones().actualizarIndicadoresTotales(df_diario, archivoSeguimiento)
            Extracciones().actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

            # Conexión con la base de datos
            FuncionalidadSQL().actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, archivoMDVR1, archivoMDVR2, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3)
            FuncionalidadSQL().actualizarInfractoresSQL(archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)
            

            ####################################
            ##### Fuera de horario laboral #####
            ####################################


            rutasLaboral = {'securitrac': archivoSecuritrac,

                        'mdvr': archivoMDVR2,

                        'ituran': archivoIturan3,

                        'ubicar': archivoUbicar2,

                        'wialon': [archivoWialon1, archivoWialon2, archivoWialon3]

                        }
            
            # Excel
            fueraHorarioLaboral = FuncionalidadExcel().fueraLaboralTodos(rutasLaboral)
            print(fueraHorarioLaboral)
            Extracciones().actualizarFueraLaboral(archivoSeguimiento, fueraHorarioLaboral)

            # SQL
            FuncionalidadSQL().sqlFueraLaboral(fueraHorarioLaboral)


            ####################################
            ######### Envío de correos #########
            ####################################


            # Enviar correo al personal de SGI.
            try:
                CorreosVehiculares().enviarCorreoPersonal()
            except Exception as e:
            
                logging.error("Ocurrió un error", exc_info=True)
            # Enviar correo específico a los conductores con excesos de velocidad.
            time.sleep(3)
            try:
                CorreosVehiculares().enviarCorreoConductor()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
            time.sleep(4)   
            # Enviar correo al personal de SGI de vehículos fuera de horario laboral.
            try:
                CorreosVehiculares().enviarCorreoLaboral()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
            ####################################
            ######## Salida del sistema ########
            ####################################


            sys.exit()


        elif hora >= 2320 and hora <= 2340:

            ########################
            ### Consulta inicial ###
            ########################

            tablaEstadosTotales = ConsultaImportante().verificarEstadosFinales()
            tablaEstadosTotales = pd.DataFrame(tablaEstadosTotales, columns=['plataforma', 'estado']).set_index('plataforma')


            ######## Todo bien inicio
            if all(ele == "Ejecutado" for ele in tablaEstadosTotales['estado'].values) == True:
                sys.exit()
            else:
                print("ya") # En caso de que no todos sean Ejecutado, no se sigue.


            ### En caso de que se siga, revisar qué plataformas tuvieron errores.
            plataformasFallidas = []
            for plataforma in tablaEstadosTotales.index: #Verifica qué plataformas tuvieron errores o no fueron ejecutadas por alguna razón.
                estado = tablaEstadosTotales.loc[plataforma]['estado']
                if estado != "Ejecutado":
                    plataformasFallidas.append(plataforma)

            listaEstadosTotales = []

            ####################################
            ###### RPA por cada plataforma #####
            ####################################


            for plataforma in plataformasFallidas: # Realiza los RPA de las plataformas fallidas
                if plataforma in plataformasFallidas and plataforma == "Ituran":
                    archivoIturan1, archivoIturan2, archivoIturan3, archivoIturan4 = RPA().ejecutarRPAIturan()
                    if archivoIturan1 == os.getcwd() + r"\archivoFicticio.csv":
                        listaEstadosTotales.append('Error')
                    else:
                        listaEstadosTotales.append('Ejecutado')
                else:
                    print("ya")
                    listaEstadosTotales.append('Ejecutado')

                if plataforma in plataformasFallidas and plataforma == "Securitrac":
                    archivoSecuritrac = RPA().ejecutarRPASecuritrac()
                    if archivoSecuritrac == os.getcwd() + r"\archivoFicticio.xls":
                        listaEstadosTotales.append('Error')
                    else:
                        listaEstadosTotales.append('Ejecutado')
                else:
                    print("ya")
                    listaEstadosTotales.append('Ejecutado')

                if plataforma in plataformasFallidas and plataforma == "MDVR":
                    archivoMDVR1,archivoMDVR2, archivoMDVR3 = RPA().ejecutarRPAMDVR()
                    if archivoMDVR1 == os.getcwd() + r"\archivoFicticio.xls":
                        listaEstadosTotales.append('Error')
                    else:
                        listaEstadosTotales.append('Ejecutado')
                else:
                    print("ya")
                    listaEstadosTotales.append('Ejecutado')

                if plataforma in plataformasFallidas and plataforma == "Ubicar":
                    archivoUbicar1,archivoUbicar2,archivoUbicar3 = RPA().ejecutarRPAUbicar()
                    if archivoUbicar1 == os.getcwd() + r"\archivoFicticio.xlsx":
                        listaEstadosTotales.append('Error')
                    else:
                        listaEstadosTotales.append('Ejecutado')
                else:
                    print("ya")
                    listaEstadosTotales.append('Ejecutado')

                if plataforma in plataformasFallidas and plataforma == "Ubicom":
                    archivoUbicom1, archivoUbicom2 = RPA().ejecutarRPAUbicom()
                    if archivoUbicom1 == os.getcwd() + r"\archivoFicticio.xls":
                        listaEstadosTotales.append('Error')
                    else:
                        listaEstadosTotales.append('Ejecutado')
                else:
                    print("ya")
                    listaEstadosTotales.append('Ejecutado')

                if plataforma in plataformasFallidas and plataforma == "Wialon":
                    archivoWialon1, archivoWialon2, archivoWialon3 = RPA().ejecutarRPAWialon()
                    if archivoWialon1 == os.getcwd() + r"\archivoFicticio.xlsx":
                        listaEstadosTotales.append('Error')
                    else:
                        listaEstadosTotales.append('Ejecutado')
                else:
                    print("ya")
                    listaEstadosTotales.append('Ejecutado')


            ####################################
            ### Verificar estados intermedios ##
            ####################################

            time.sleep(5)
            # Si alguna plataforma falló, pero se arregló, se ejecutará el archivo. Si no pasa esto, se sale.
            if all(ele == "Ejecutado" for ele in listaEstadosTotales) == False:
                print(listaEstadosTotales)
                sys.exit() # En caso de que no todos sean Ejecutado, no se sigue.

            print("sigue")

            
            ####################################
            ####### Creación de informes #######
            ####################################


            archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

            # Actualización de seguimiento
            df_exist = Extracciones().crear_excel(archivoMDVR1,archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)

            # Actualización de infractores
            Extracciones().actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

            # Actualización del odómetro
            Extracciones().actualizarOdom(archivoSeguimiento, archivoIturan4, archivoUbicar1)

            # Actualización de indicadores
            df_diario = Extracciones().dfDiario(df_exist)
            Extracciones().actualizarIndicadoresTotales(df_diario, archivoSeguimiento)
            Extracciones().actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

            # Conexión con la base de datos
            FuncionalidadSQL().actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, archivoMDVR1, archivoMDVR2, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3)
            FuncionalidadSQL().actualizarInfractoresSQL(archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)
            

            ####################################
            ##### Fuera de horario laboral #####
            ####################################

            rutasLaboral = {'securitrac': archivoSecuritrac,

                        'mdvr': archivoMDVR2,

                        'ituran': archivoIturan3,

                        'ubicar': archivoUbicar2,

                        'wialon': [archivoWialon1, archivoWialon2, archivoWialon3]

                        }
            
            # Excel
            fueraHorarioLaboral = FuncionalidadExcel().fueraLaboralTodos(rutasLaboral)
            Extracciones().actualizarFueraLaboral(archivoSeguimiento, fueraHorarioLaboral)

            # SQL
            FuncionalidadSQL().sqlFueraLaboral(fueraHorarioLaboral)

            ####################################
            ######### Envío de correos #########
            ####################################


            # Enviar correo al personal de SGI.
            try:
                CorreosVehiculares().enviarCorreoPersonal()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
            time.sleep(3)
            # Enviar correo específico a los conductores con excesos de velocidad.
            try:
                CorreosVehiculares().enviarCorreoConductor()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
            time.sleep(4)
            # Enviar correo al personal de SGI de vehículos fuera de horario laboral.
            try:
                CorreosVehiculares().enviarCorreoLaboral()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)

            ####################################
            ######## Salida del sistema ########
            ####################################


            sys.exit()
        

        elif hora >= 2340:

            ########################
            ### Consulta inicial ###
            ########################

            tablaEstadosTotales = ConsultaImportante().verificarEstadosFinales()
            tablaEstadosTotales = pd.DataFrame(tablaEstadosTotales, columns=['plataforma', 'estado']).set_index('plataforma')


            ######## Todo bien inicio
            if all(ele == "Ejecutado" for ele in tablaEstadosTotales['estado'].values) == True:
                TratadorArchivos().eliminarArchivosOutput()
                ConsultaImportante().actualizarTablaEstados()
                sys.exit() # En caso de que no todos sean Ejecutado, no se sigue.


            ### En caso de que se siga, revisar qué plataformas tuvieron errores.
            plataformasFallidas = []
            for plataforma in tablaEstadosTotales.index: #Verifica qué plataformas definitivamente tuvieron errores.
                estado = tablaEstadosTotales.loc[plataforma]['estado']
                if estado != "Ejecutado":
                    plataformasFallidas.append(plataforma)

            listaEstadosTotales = []

            ####################################
            ###### RPA por cada plataforma #####
            ####################################


            for plataforma in plataformasFallidas: # Realiza los RPA de las plataformas fallidas
                if plataforma in plataformasFallidas and plataforma == "Ituran":
                    archivoIturan1, archivoIturan2, archivoIturan3, archivoIturan4 = RPA().ejecutarRPAIturan()
                    if archivoIturan1 == os.getcwd() + r"\archivoFicticio.csv":
                        listaEstadosTotales.append('ituranError')
                    else:
                        listaEstadosTotales.append('ituranEjecutado')
                else:
                    print("ya")

                if plataforma in plataformasFallidas and plataforma == "Securitrac":
                    archivoSecuritrac = RPA().ejecutarRPASecuritrac()
                    if archivoSecuritrac == os.getcwd() + r"\archivoFicticio.xls":
                        listaEstadosTotales.append('securitacError')
                    else:
                        listaEstadosTotales.append('securitracEjecutado')
                else:
                    print("ya")

                if plataforma in plataformasFallidas and plataforma == "MDVR":
                    archivoMDVR1,archivoMDVR2, archivoMDVR3 = RPA().ejecutarRPAMDVR()
                    if archivoMDVR1 == os.getcwd() + r"\archivoFicticio.xls":
                        listaEstadosTotales.append('MDVRError')
                    else:
                        listaEstadosTotales.append('MDVREjecutado')
                else:
                    print("ya")

                if plataforma in plataformasFallidas and plataforma == "Ubicar":
                    archivoUbicar1,archivoUbicar2,archivoUbicar3 = RPA().ejecutarRPAUbicar()
                    if archivoUbicar1 == os.getcwd() + r"\archivoFicticio.xlsx":
                        listaEstadosTotales.append('ubicarError')
                    else:
                        listaEstadosTotales.append('ubicarEjecutado')
                else:
                    print("ya")

                if plataforma in plataformasFallidas and plataforma == "Ubicom":
                    archivoUbicom1, archivoUbicom2 = RPA().ejecutarRPAUbicom()
                    if archivoUbicom1 == os.getcwd() + r"\archivoFicticio.xls":
                        listaEstadosTotales.append('ubicomError')
                    else:
                        listaEstadosTotales.append('ubicomEjecutado')
                else:
                    print("ya")

                if plataforma in plataformasFallidas and plataforma == "Wialon":
                    archivoWialon1, archivoWialon2, archivoWialon3 = RPA().ejecutarRPAWialon()
                    if archivoWialon1 == os.getcwd() + r"\archivoFicticio.xlsx":
                        listaEstadosTotales.append('wialonError')
                    else:
                        listaEstadosTotales.append('wialonEjecutado')
                else:
                    print("ya")

            ####################################
            ##### Verificar estados finales ####
            ####################################
            Correos = CorreosVehiculares()
            Tratador = TratadorArchivos()
            # Si alguna plataforma falló definitivamente, aparecerá aquí y se sigue con la ejecución normal.
            for estado in listaEstadosTotales: #Verifica qué plataformas definitivamente tuvieron errores.
                if estado == "ituranError":
                    ConsultaImportante().registrarError("Ituran")
                    Correos.enviarCorreoPlataforma("Ituran")
                    Tratador.crearDirectorioError('Ituran')
                if estado == "securitracError":
                    ConsultaImportante().registrarError("Securitrac")
                    Correos.enviarCorreoPlataforma("Securitrac")
                    Tratador.crearDirectorioError('Securitrac')
                if estado == "MDVRError":
                    ConsultaImportante().registrarError("MDVR")
                    Correos.enviarCorreoPlataforma("MDVR")
                    Tratador.crearDirectorioError('MDVR')
                if estado == "UbicarError":
                    ConsultaImportante().registrarError("Ubicar")
                    Correos.enviarCorreoPlataforma("Ubicar")
                    Tratador.crearDirectorioError('Ubicar')
                if estado == "UbicomError":
                    ConsultaImportante().registrarError("Ubicom")
                    Correos.enviarCorreoPlataforma("Ubicom")
                    Tratador.crearDirectorioError('Ubicom')
                if estado == "wialonError":
                    ConsultaImportante().registrarError("Wialon")
                    Correos.enviarCorreoPlataforma("Wialon")
                    Tratador.crearDirectorioError('Wialon')
                else:
                    print("hecho")

            ####################################
            ####### Creación de informes #######
            ####################################


            archivoSeguimiento = os.getcwd() + "\\seguimiento.xlsx"

            # Actualización de seguimiento
            df_exist = Extracciones().crear_excel(archivoMDVR1,archivoMDVR3, archivoIturan1, archivoIturan2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSeguimiento)

            # Actualización de infractores
            Extracciones().actualizarInfractores(archivoSeguimiento, archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)

            # Actualización del odómetro
            Extracciones().actualizarOdom(archivoSeguimiento, archivoIturan4, archivoUbicar1)

            # Actualización de indicadores
            df_diario = Extracciones().dfDiario(df_exist)
            Extracciones().actualizarIndicadoresTotales(df_diario, archivoSeguimiento)
            Extracciones().actualizarIndicadores(df_diario, df_exist, archivoSeguimiento)

            # Conexión con la base de datos
            FuncionalidadSQL().actualizarSeguimientoSQL(archivoIturan1, archivoIturan2, archivoMDVR1, archivoMDVR2, archivoUbicar1, archivoUbicar2, archivoUbicom1, archivoUbicom2, archivoSecuritrac, archivoWialon1, archivoWialon2, archivoWialon3)
            FuncionalidadSQL().actualizarInfractoresSQL(archivoIturan2, archivoMDVR3, archivoUbicar3, archivoWialon1, archivoWialon2, archivoWialon3, archivoSecuritrac)
            

            ####################################
            ##### Fuera de horario laboral #####
            ####################################

            rutasLaboral = {'securitrac': archivoSecuritrac,

                        'mdvr': archivoMDVR2,

                        'ituran': archivoIturan3,

                        'ubicar': archivoUbicar2,

                        'wialon': [archivoWialon1, archivoWialon2, archivoWialon3]

                        }
            
            # Excel
            fueraHorarioLaboral = FuncionalidadExcel().fueraLaboralTodos(rutasLaboral)
            Extracciones().actualizarFueraLaboral(archivoSeguimiento, fueraHorarioLaboral)

            # SQL
            FuncionalidadSQL().sqlFueraLaboral(fueraHorarioLaboral)

            ####################################
            ######### Envío de correos #########
            ####################################


            # Enviar correo al personal de SGI.
            try:
                CorreosVehiculares().enviarCorreoPersonal()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
            time.sleep(2)
            # Enviar correo específico a los conductores con excesos de velocidad.
            try:
                CorreosVehiculares().enviarCorreoConductor()
            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
                time.sleep(7)
            # Enviar correo al personal de SGI de vehículos fuera de horario laboral.
            try:
                CorreosVehiculares().enviarCorreoLaboral()

            except Exception as e:
    
                logging.error("Ocurrió un error", exc_info=True)
            ####################################
            ######### Borrado y salida #########
            ####################################


            # Eliminar las carpetas del output ya que se tiene toda la información.
            print("Eliminando archivos")
            time.sleep(5)
            TratadorArchivos().eliminarArchivosOutput()

            # Actualización de la tabla de estados.
            ConsultaImportante().actualizarTablaEstados()


            # Salida del sistema.
            sys.exit()
        

        else:
            print("No tiene sentido llegar aquí.")

    except Exception as e:
    
        logging.error("Ocurrió un error", exc_info=True)
    

if __name__=='__main__':
    logging.basicConfig(level=logging.ERROR, format='%(asctime)s %(levelname)s %(message)s', filename='error.log')
    main()
    



