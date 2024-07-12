import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from pathlib import Path
from db.conexionDB import conexionDB
import datetime
import pandas as pd
from selenium import webdriver


####################################
### Enviar correo a personal SGI ###
####################################


def enviarCorreoPersonal():
    """
    Realiza el proceso del envío del correo al personal de SGI interesado.
    """

    ########### Conexión con la base de datos.

    # Tabla del correo.
    conexionBaseCorreos = conexionDB().establecerConexion()
    if conexionBaseCorreos:
        cursor = conexionBaseCorreos.cursor()
    else:
        print("Error.")
    
    #Consulta de los correos necesarios para el correo.
    cursor.execute("select * from vehiculos.infractorVehicular where numeroExcesosVelocidad >0")
    tablaExcesos = cursor.fetchall()
    cursor.execute("select * from vehiculos.correoVehicular")
    tablaCorreos = cursor.fetchall()

    #Desconectar BD
    conexionDB().cerrarConexion()

    ##########

    # Modificaciones iniciales a los datos de las consultas.
    tablaExcesos = pd.DataFrame(tablaExcesos, columns=['eliminar','Conductor', 'Correo', 'Número de excesos de velocidad']).drop(['eliminar','Correo'],axis=1)
    tablaCorreos = pd.DataFrame(tablaCorreos,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')
    correoReceptor = tablaCorreos['correo'].dropna().tolist()
    correoCopia = tablaCorreos['correoCopia'].dropna().tolist()
    # Datos sobre el correo.
    correoEmisor = 'notificaciones.sgi@appsgi.com.co'
    correoDestinatarios = correoReceptor + correoCopia


    # Texto del correo.
    correoTexto = f"""Buenos días. Espero que se encuentre bien.
    
    Mediante el presente correo puede encontrar el informe vehicular actualizado hasta el día de hoy.
    En este, podrá encontrar información como el número y duración de los excesos de velocidad, el kilometraje diario y total del vehículo, o el número de desplazamientos de cada vehículo.
    Asimismo, mediante el presente correo puede encontrar una tabla con el número de excesos de velocidad por vehículo, con su respectivo nombre del conductor (en caso de que sea fijo) y placa.
    
    {tablaExcesos}
    
    Atentamente,
    Departamento de Tecnología y desarrollo, SGI SAS"""
    
    correoAsunto = f'Informe de seguimiento a vehículos del día {datetime.date.today()}'
    plataformasFinalRuta = os.getcwd() + '\\plataformasFinal.xlsx'

    mensajeCorreo = MIMEMultipart()
    mensajeCorreo['From'] = f"{Header('Notificacion SGI', 'utf-8')} <{correoEmisor}>"
    mensajeCorreo['To'] = ", ".join(correoReceptor)
    mensajeCorreo['Cc'] = ", ".join(correoCopia)
    mensajeCorreo['Subject'] = correoAsunto
    mensajeCorreo.attach(MIMEText(correoTexto, 'plain'))
    with open(plataformasFinalRuta, "rb") as ruta:
        r=MIMEApplication(ruta.read(), Name="plataformasFinal.xlsx")
        r.set_payload(ruta.read())
    encoders.encode_base64(r)
    r.add_header('Content-Disposition','attachment; filename={}'.format(Path(plataformasFinalRuta).name))
    mensajeCorreo.attach(r)

    # Inicializar el correo y enviar.
    servidorCorreo = smtplib.SMTP('smtp.hostinger.com', 587)
    servidorCorreo.starttls()
    servidorCorreo.login(correoEmisor, '$f~Pu$9zUIu)%=3')
    servidorCorreo.sendmail(correoEmisor, correoDestinatarios, mensajeCorreo.as_string())
    servidorCorreo.quit()


####################################
##### Eliminar archivos del día ####
####################################

def eliminarArchivosOutput():
    """
    Elimina los archivos que aparecen en las carpetas de Output de cada RPA de cada plataforma.
    """
    for folder in os.listdir():
        if "output" in folder:
            shutil.rmtree(folder)


####################################
#### Enviar correo al conductor ####
####################################


def enviarCorreoConductor():
    """
    Realiza el proceso del envío del correo a los conductores que tuvieron excesos de velocidad.
    """

    ########### Conexión con la base de datos.

    # Tabla del correo.
    conexionBaseCorreos = conexionDB().establecerConexion()
    if conexionBaseCorreos:
        cursor = conexionBaseCorreos.cursor()
    else:
        print("Error.")
    
    #Consulta de los correos necesarios para el correo.
    cursor.execute("select * from vehiculos.infractorVehicular where numeroExcesosVelocidad >0")
    tablaExcesos2 = cursor.fetchall() #Obtener todos los resultados
    
    #Desconectar BD
    conexionDB().cerrarConexion()

    ##########

    # Ajustes adicionales a la tabla de excesos 2.
    tablaExcesos2 = pd.DataFrame(tablaExcesos2, columns=['eliminar','Conductor', 'Correo', 'Número de excesos de velocidad','Placa']).drop(['eliminar','Placa'],axis=1)
    listaConductores = tablaExcesos2['Conductor'].tolist()

    #### Loop para realizar el envío del correo.
    for conductorVehicular in listaConductores:
        tablaExcesos3 = tablaExcesos2[tablaExcesos2['Conductor'] == conductorVehicular]
        tablaExcesos3 = tablaExcesos3.set_index('Conductor')

        # Datos sobre el correo.
        correoEmisor = 'notificaciones.sgi@appsgi.com.co'
        correoReceptor = tablaExcesos3.loc[conductorVehicular]['Correo']

        # Texto del correo.
        correoTexto = f"""Buenos días. Espero que se encuentre bien.
        
        Mediante el presente correo puede encontrar los excesos de velocidad que usted tuvo en el día.
        Esta información le puede ayudar a mejorar sus hábitos de conducción y, de esta manera, evitar posibles siniestros viales.
        
        Conductor: {tablaExcesos3.reset_index().iloc[0]['Conductor']}
        Número de excesos de velocidad: {tablaExcesos3.loc[conductorVehicular]['Número de excesos de velocidad']}
        
        Atentamente,
        Departamento de Tecnología y desarrollo, SGI SAS"""
        
        correoAsunto = f'Informe de conducción individual de {tablaExcesos3.reset_index().iloc[0]['Conductor']} para el {datetime.date.today()}'

        mensajeCorreo = MIMEMultipart()
        mensajeCorreo['From'] = f"{Header('Notificaciones SGI', 'utf-8')} <{correoEmisor}>"
        mensajeCorreo['To'] = correoReceptor
        mensajeCorreo['Subject'] = correoAsunto
        mensajeCorreo.attach(MIMEText(correoTexto, 'plain'))

        # Inicializar el correo y enviar.
        servidorCorreo = smtplib.SMTP('smtp.hostinger.com', 587)
        servidorCorreo.starttls()
        servidorCorreo.login(correoEmisor, '$f~Pu$9zUIu)%=3')
        servidorCorreo.sendmail(correoEmisor, correoReceptor, mensajeCorreo.as_string())
        servidorCorreo.quit()


####################################
###### Definir ruta navegador ######
####################################

class Navegador():
    def rutaNavegador(plataforma):
        opcionesNavegador = webdriver.ChromeOptions()
        carpetaOutput = r"\output" + plataforma
        lugarDescargas = os.getcwd() + carpetaOutput
        if not os.path.exists(lugarDescargas):
            os.makedirs(lugarDescargas)

        opcionDescarga = {
            "download.default_directory": lugarDescargas,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        }