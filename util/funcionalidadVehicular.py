import os
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from pathlib import Path

import pandas as pd


####################################
### Enviar correo a personal SGI ###
####################################


def enviarCorreo():
    """
    Realiza el proceso del envío del correo al personal de SGI interesado.
    """


    correoEmisor = 'notificaciones.sgi@appsgi.com.co'
    correoReceptor = 'pab.aoarc@gmail.com'
    correoCopia = ['p.ojeda@uniandes.edu.co']
    correoDestinatarios = [correoReceptor] + correoCopia
    
    ########## CAMBIAR POR TABLA DE SQL.
    tablaExcesos = pd.read_excel("plataformasFinal.xlsx",index_col='Placa')
    ########## CAMBIAR POR TABLA DE SQL.

    correoTexto = f"""Buenos días. Espero que se encuentre bien.
    
    Mediante el presente correo puede encontrar el informe vehicular actualizado hasta el día de hoy.
    En este, podrá encontrar información como el número y duración de los excesos de velocidad, el kilometraje diario y total del vehículo, o el número de desplazamientos de cada vehículo.
    Asimismo, mediante el presente correo puede encontrar una tabla con el número de excesos de velocidad por vehículo, con su respectivo nombre del conductor (en caso de que sea fijo) y placa.
    
    {tablaExcesos}
    
    Atentamente,
    Departamento de Tecnología y desarrollo, SGI SAS"""
    
    correoAsunto = 'pruebaPRUEBA RPA VEHÍCULOS'
    plataformasFinalRuta = os.getcwd() + '\\plataformasFinal.xlsx'


    mensajeCorreo = MIMEMultipart()
    mensajeCorreo['From'] = f"{Header('Notificacion SGI', 'utf-8')} <{correoEmisor}>"
    mensajeCorreo['To'] = correoReceptor
    mensajeCorreo['Cc'] = ", ".join(correoCopia)
    mensajeCorreo['Subject'] = correoAsunto
    mensajeCorreo.attach(MIMEText(correoTexto, 'plain'))
    with open(plataformasFinalRuta, "rb") as ruta:
        r=MIMEApplication(ruta.read(), Name="plataformasFinal.xlsx")
        r.set_payload(ruta.read())
    encoders.encode_base64(r)
    r.add_header('Content-Disposition','attachment; filename={}'.format(Path(plataformasFinalRuta).name))
    mensajeCorreo.attach(r)

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
