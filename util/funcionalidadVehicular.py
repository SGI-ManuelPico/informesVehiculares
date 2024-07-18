import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from db.conexionDB import conexionDB
import datetime
import pandas as pd
import textwrap
from pretty_html_table import build_table


####################################
### Enviar correo a personal SGI ###
####################################


def enviarCorreoPersonal():
    """
    Realiza el proceso del envío del correo al personal de SGI interesado.
    """

    ######### Tabla del correo.
    conexionBaseCorreos = conexionDB().establecerConexion()
    if conexionBaseCorreos:
        cursor = conexionBaseCorreos.cursor()
    else:
        print("Error.")

    #Consulta de los correos necesarios para el correo.
    cursor.execute("SELECT placa, tiempoExceso, VelocidadMaxima, conductor FROM vehiculos.infractor where date(fecha) like curdate();")
    tablaExcesos = cursor.fetchall()
    cursor.execute("select * from vehiculos.plataformasVehiculares")
    tablaCorreos = cursor.fetchall()

    #Desconectar BD
    conexionDB().cerrarConexion()

    ##########

    # Modificaciones iniciales a los datos de las consultas.
    tablaCorreos = pd.DataFrame(tablaCorreos,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')
    tablaExcesos = pd.DataFrame(tablaExcesos, columns=['Placa', 'Duración', 'Velocidad', 'Conductor'])
   
    correoReceptor = tablaCorreos['correo'].dropna().tolist()
    correoCopia = tablaCorreos['correoCopia'].dropna().tolist()
    # Datos sobre el correo.
    correoEmisor = 'notificaciones.sgi@appsgi.com.co'
    correoDestinatarios = correoReceptor + correoCopia

    # Texto del correo.
    correoTexto = f"""
    <p>Buenos d&iacute;as. Espero que se encuentre bien.</p>

    <p>Mediante el presente correo puede encontrar el informe vehicular actualizado hasta el d&iacute;a de hoy.<br>
    En este, podr&aacute; encontrar informaci&oacute;n como el n&uacute;mero y duraci&oacute;n de los excesos de velocidad, el kilometraje diario y total del veh&iacute;culo, o el n&uacute;mero de desplazamientos de cada veh&iacute;culo.<br>
    Asimismo, mediante el presente correo puede encontrar una tabla con el n&uacute;mero de excesos de velocidad por veh&iacute;culo, con su respectivo nombre del conductor (en caso de que sea fijo) y placa.</p>

    {build_table(tablaExcesos, 'green_light')}

    <p>Atentamente,<br>
    Departamento de tecnología y desarrollo, SGI SAS</p>"""

    correoTexto = textwrap.dedent(correoTexto)
    correoAsunto = f'Informe de seguimiento a vehículos del día {datetime.date.today()}'

    mensajeCorreo = MIMEMultipart()
    mensajeCorreo['From'] = f"{Header('Notificaciones SGI', 'utf-8')} <{correoEmisor}>"
    mensajeCorreo['To'] = ", ".join(correoReceptor)
    mensajeCorreo['Cc'] = ", ".join(correoCopia)
    mensajeCorreo['Subject'] = correoAsunto
    mensajeCorreo.attach(MIMEText(correoTexto, 'html'))
    
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("seguimiento.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="seguimiento.xlsx"')
    mensajeCorreo.attach(part)

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

    ######### Tabla del correo.
    conexionBaseCorreos = conexionDB().establecerConexion()
    if conexionBaseCorreos:
        cursor = conexionBaseCorreos.cursor()
    else:
        print("Error.")

    #Consulta de los correos necesarios para el correo.
    cursor.execute("SELECT placa, tiempoExceso, conductor FROM vehiculos.infractor where date(fecha) like curdate();")
    tablaExcesos = cursor.fetchall()
    cursor.execute("select * from vehiculos.plataformasVehiculares")
    tablaCorreos2 = cursor.fetchall()

    #Desconectar BD
    conexionDB().cerrarConexion()

    ##########

    # Modificaciones iniciales a los datos de las consultas.
    tablaCorreos2 = pd.DataFrame(tablaCorreos2,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')

    tablaExcesos2 = pd.DataFrame(tablaExcesos, columns=['Placa', 'Duración de excesos de velocidad', 'Conductor'])
    tablaExcesos2['Número de excesos de velocidad'] = 1
    tablaExcesos2['correoCopia'] = tablaCorreos2.iloc[0]['correo']
    tablaExcesos2['correo'] = tablaCorreos2.iloc[1]['correoCopia'] ########################### CAMBIAR AL CORREO DEL CONDUCTOR QUE APARECERÍA CON LA BASE DE DATOS ORIGINAL DE INGFRACTORES
    tablaExcesos2 = tablaExcesos2.groupby('Placa', as_index=False).agg({'Duración de excesos de velocidad': 'sum', 'Conductor':'first', 'correo': 'first', 'correoCopia' : 'first', 'Número de excesos de velocidad' : 'sum'})


    listaConductores = tablaExcesos2['Conductor'].tolist()

    #### Loop para realizar el envío del correo.
    for conductorVehicular in listaConductores:
        tablaExcesos3 = tablaExcesos2[tablaExcesos2['Conductor'] == conductorVehicular]
        tablaExcesos3 = tablaExcesos3.set_index('Conductor')

        # Datos sobre el correo.
        correoEmisor = 'notificaciones.sgi@appsgi.com.co'
        correoReceptor = tablaExcesos3.loc[conductorVehicular]['correo']
        correoCopia = tablaExcesos3.loc[conductorVehicular]['correoCopia']
        correoDestinatarios = [correoReceptor] + [correoCopia]
        correoAsunto = f'Informe de conducción individual de {tablaExcesos3.reset_index().iloc[0]['Conductor']} para el {datetime.date.today()}'

        # Texto del correo.
        correoTexto = f"""
        Buenos días. Espero que se encuentre bien.
        
        Mediante el presente correo puede encontrar los excesos de velocidad que usted tuvo en el día. Esta información le puede ayudar a mejorar sus hábitos de conducción y, de esta manera, evitar posibles siniestros viales.
        
        Conductor: {tablaExcesos3.reset_index().iloc[0]['Conductor']}
        Número de excesos de velocidad: {tablaExcesos3.loc[conductorVehicular]['Número de excesos de velocidad']}
        Placa del vehículo que maneja: {tablaExcesos3.loc[conductorVehicular]['Placa']}
        """

        if tablaExcesos3.loc[conductorVehicular]['Duración de excesos de velocidad'] >300:
            correoTexto2 = f"""
            Adicionalmente, se encontró que sus excesos de velocidad acumularon más de 5 minutos en total. Específicamente, su duración total en exceso fue de {tablaExcesos3.loc[conductorVehicular]['Duración de excesos de velocidad']} segundos. Esta información le puede ser de vital importancia para evitar situaciones que le puedan colocar en un riesgo importante para su vida.

            Atentamente,
            Departamento de Tecnología y desarrollo, SGI SAS
            """
        else:
            correoTexto2 = f"""
            
            Atentamente,
            Departamento de Tecnología y desarrollo, SGI SAS
            """
        
        # Para formatear el texto de manera correcta.
        correoTexto2 = textwrap.dedent(correoTexto2)
        correoTexto = textwrap.dedent(correoTexto)
        correoTexto = correoTexto + correoTexto2

        # Creación del correo.
        mensajeCorreo = MIMEMultipart()
        mensajeCorreo['From'] = f"{Header('Notificaciones SGI', 'utf-8')} <{correoEmisor}>"
        mensajeCorreo['To'] = correoReceptor
        mensajeCorreo['Cc'] = correoCopia
        mensajeCorreo['Subject'] = correoAsunto
        mensajeCorreo.attach(MIMEText(correoTexto, 'plain'))


        # Inicializar el correo y enviar.
        servidorCorreo = smtplib.SMTP('smtp.hostinger.com', 587)
        servidorCorreo.starttls()
        servidorCorreo.login(correoEmisor, '$f~Pu$9zUIu)%=3')
        servidorCorreo.sendmail(correoEmisor, correoDestinatarios, mensajeCorreo.as_string())
        servidorCorreo.quit()


####################################
## Correo plataforma disfuncional ##
####################################


def enviarCorreoPlataforma(plataforma):
    """
    Realiza el proceso del envío del correo a los interesados en caso de que una plataforma no haya funcionado.
    """

    ######### Tabla del correo.
    conexionBaseCorreos = conexionDB().establecerConexion()
    if conexionBaseCorreos:
        cursor = conexionBaseCorreos.cursor()
    else:
        print("Error.")

    #Consulta de los correos necesarios para el correo.
    cursor.execute("select * from vehiculos.plataformasVehiculares")
    tablaCorreos2 = cursor.fetchall()

    #Desconectar BD
    conexionDB().cerrarConexion()

    ########## Envío del correo

    # Modificaciones iniciales a los datos de las consultas.
    tablaCorreos2 = pd.DataFrame(tablaCorreos2,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')

    # Datos sobre el correo.
    correoEmisor = 'notificaciones.sgi@appsgi.com.co'
    correoReceptor = tablaCorreos2['correo'].dropna().tolist()
    correoCopia = tablaCorreos2['correoCopia'].dropna().tolist()
    correoDestinatarios = correoReceptor + correoCopia
    correoAsunto = f'Notificación de errores para la plataforma {plataforma} durante el día {datetime.date.today()}'

    # Texto del correo.
    correoTexto = f"""
    Buenos días. Espero que se encuentre bien.
    
    Mediante el presente correo se le informa que la plataforma {plataforma} tuvo errores durante su ejecución. Le invito a investigar más al respecto y, en cualquiera de los casos, el departamento de tecnología y desarrollo estará atento a sus inquietudes.

    Es importante aclarar que el informe dejará los valores asociados a los vehículos de {plataforma} como "0" y esto deberá ser corregido manualmente.

    Atentamente,
    Departamento de Tecnología y desarrollo, SGI SAS
    """
    
    # Para formatear el texto de manera correcta.
    correoTexto = textwrap.dedent(correoTexto)

    # Creación del correo.
    mensajeCorreo = MIMEMultipart()
    mensajeCorreo['From'] = f"{Header('Notificaciones SGI', 'utf-8')} <{correoEmisor}>"
    mensajeCorreo['To'] = ", ".join(correoReceptor)
    mensajeCorreo['Cc'] = ", ".join(correoCopia)
    mensajeCorreo['Subject'] = correoAsunto
    mensajeCorreo.attach(MIMEText(correoTexto, 'plain'))


    # Inicializar el correo y enviar.
    servidorCorreo = smtplib.SMTP('smtp.hostinger.com', 587)
    servidorCorreo.starttls()
    servidorCorreo.login(correoEmisor, '$f~Pu$9zUIu)%=3')
    servidorCorreo.sendmail(correoEmisor, correoDestinatarios, mensajeCorreo.as_string())
    servidorCorreo.quit()