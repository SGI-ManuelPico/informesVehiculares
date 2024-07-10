import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
from pathlib import Path

correoEmisor = 'notificaciones.sgi@appsgi.com.co'
correoReceptor = 'pab.aoarc@gmail.com'
correo = [correoReceptor]
correoTexto = 'pruebaPRUEBA RPA VEHÍCULOS PRUEBA PARA VER SI FUNCIONA.'
correoAsunto = 'pruebaPRUEBA RPA VEHÍCULOS'
plataformasFinalRuta = os.getcwd() + '\\plataformasFinal.xlsx'


mensajeCorreo = MIMEMultipart()
mensajeCorreo['From'] = f"{Header('Notificacion SGI', 'utf-8')} <{correoEmisor}>"
mensajeCorreo['To'] = correoReceptor
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
servidorCorreo.sendmail(correoEmisor, correoReceptor, mensajeCorreo.as_string())
servidorCorreo.quit()