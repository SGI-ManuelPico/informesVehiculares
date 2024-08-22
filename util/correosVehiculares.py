import smtplib, time
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.header import Header
from email import encoders
import datetime
import pandas as pd
import textwrap
from pretty_html_table import build_table
from db.consultasImportantes import ConsultaImportante



class CorreosVehiculares:
    def __init__(self):
        pass

    ####################################
    ### Enviar correo a personal SGI ###
    ####################################


    def enviarCorreoPersonal(self):
        """
        Realiza el proceso del envío del correo al personal de SGI interesado.
        """

        self.tablaExcesos, self.tablaCorreos = ConsultaImportante().tablaCorreoPersonal()

        # Modificaciones iniciales a los datos de las consultas.
        self.tablaCorreos = pd.DataFrame(self.tablaCorreos,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')
        self.tablaExcesos = pd.DataFrame(self.tablaExcesos, columns=['Placa', 'Duración', 'Velocidad', 'Conductor'])
    
        correoReceptor = self.tablaCorreos['correo'].dropna().tolist()
        correoCopia = self.tablaCorreos['correoCopia'].dropna().tolist()
        # Datos sobre el correo.
        correoEmisor = 'notificaciones.sgi@appsgi.com.co'
        correoDestinatarios = correoReceptor + correoCopia

        # Texto del correo.
        correoTexto = f"""
        <p>Buenos d&iacute;as. Espero que se encuentre bien.</p>

        <p>Mediante el presente correo puede encontrar el informe vehicular actualizado hasta el d&iacute;a de hoy.<br>
        En este, podr&aacute; encontrar informaci&oacute;n como el n&uacute;mero y duraci&oacute;n de los excesos de velocidad, el kilometraje diario y total del veh&iacute;culo, o el n&uacute;mero de desplazamientos de cada veh&iacute;culo.<br>
        Asimismo, puede encontrar una tabla con el n&uacute;mero de excesos de velocidad por veh&iacute;culo, con su respectivo conductor y placa.</p>

        {build_table(self.tablaExcesos, 'green_light')}

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
    #### Enviar correo al conductor ####
    ####################################


    def enviarCorreoConductor(self):
        """
        Realiza el proceso del envío del correo a los conductores que tuvieron excesos de velocidad.
        """

        self.tablaExcesos, self.tablaCorreos = ConsultaImportante().tablaCorreoPersonal()

        # Modificaciones iniciales a los datos de las consultas.
        self.tablaCorreos = pd.DataFrame(self.tablaCorreos,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')
        self.tablaExcesos = pd.DataFrame(self.tablaExcesos, columns=['Placa', 'Duración', 'Velocidad', 'Conductor'])
        self.tablaExcesos2 = self.tablaExcesos
        self.tablaExcesos2['Número'] = 1
        self.tablaExcesos2 = self.tablaExcesos2.drop(columns='Velocidad').groupby('Placa', as_index=False).agg({'Duración': 'sum', 'Conductor':'first', 'Número' : 'sum'})


        listaConductores = self.tablaExcesos2['Conductor'].tolist()


        #### Loop para realizar el envío del correo.
        for conductorVehicular in listaConductores:
            time.sleep(5)
            self.tablaExcesos3 = self.tablaExcesos2[self.tablaExcesos2['Conductor'] == conductorVehicular]
            self.tablaExcesos3 = self.tablaExcesos3.set_index('Conductor')

            # Datos sobre el correo.
            correoEmisor = 'notificaciones.sgi@appsgi.com.co'
            correoReceptor = self.tablaCorreos['correoCopia'].dropna().tolist()
            correoCopia = self.tablaCorreos['correo'].dropna().tolist()
            correoDestinatarios = [correoReceptor] + [correoCopia]
            correoAsunto = f"Informe de conducción individual de {self.tablaExcesos3.reset_index().iloc[0]['Conductor']} para el {datetime.date.today()}"
            
            # Texto del correo.
            correoTexto = f"""
            <p>Buenos d&iacute;as. Espero que se encuentre bien.</p>

            <p>Mediante el presente correo puede encontrar los excesos de velocidad que usted tuvo en el d&iacute;a. Esta informaci&oacute;n le puede ayudar a mejorar sus h&aacute;bitos de conducci&oacute;n y, de esta manera, evitar posibles siniestros viales.</p>

            <p>Conductor: {self.tablaExcesos3.reset_index().iloc[0]['Conductor']}<br>
            N&uacute;mero de excesos de velocidad: {self.tablaExcesos3.loc[conductorVehicular]['Número']}<br>
            Placa del veh&iacute;culo que maneja: {self.tablaExcesos3.loc[conductorVehicular]['Placa']}</p>
            """

            self.tablaExcesos = self.tablaExcesos[['Placa', 'Duración', 'Velocidad', 'Conductor']]

            if self.tablaExcesos3.loc[conductorVehicular]['Duración'] >300:
                correoTexto2 = f"""
                <p>Adicionalmente, se encontr&oacute; que sus excesos de velocidad acumularon m&aacute;s de 5 minutos en total. Espec&iacute;ficamente, su duraci&oacute;n total en exceso fue de {self.tablaExcesos3.loc[conductorVehicular]['Duración']} segundos. Esta informaci&oacute;n le puede ser de vital importancia para evitar situaciones que le puedan colocar en un riesgo importante para su vida.</p>

                <p>Por otro lado, tambi&eacute;n se le env&iacute;a su registro de n&uacute;mero de excesos de velocidad realizados durante el d&iacute;a.</p>


                {build_table(self.tablaExcesos[self.tablaExcesos['Conductor'] == conductorVehicular], 'green_light')}

                Atentamente,
                Departamento de Tecnología y desarrollo, SGI SAS
                """
            else:
                correoTexto2 = f"""
                <p>Por otro lado, a trav&eacute;s de este medio se le env&iacute;a su registro de n&uacute;mero de excesos de velocidad realizados durante el d&iacute;a.</p>

                {build_table(self.tablaExcesos[self.tablaExcesos['Conductor'] == conductorVehicular], 'green_light')}

                <p>Atentamente,<br>
                Departamento de Tecnolog&iacute;a y desarrollo, SGI SAS</p>

                """
            
            # Para formatear el texto de manera correcta.
            correoTexto2 = textwrap.dedent(correoTexto2)
            correoTexto = textwrap.dedent(correoTexto)
            correoTexto = correoTexto + correoTexto2

            # Creación del correo.
            mensajeCorreo = MIMEMultipart()
            mensajeCorreo['From'] = f"{Header('Notificaciones SGI', 'utf-8')} <{correoEmisor}>"
            mensajeCorreo['To'] = ", ".join(correoReceptor)
            mensajeCorreo['Cc'] = ", ".join(correoCopia)
            mensajeCorreo['Subject'] = correoAsunto
            mensajeCorreo.attach(MIMEText(correoTexto, 'html'))


            # Inicializar el correo y enviar.
            servidorCorreo = smtplib.SMTP('smtp.hostinger.com', 587)
            servidorCorreo.starttls()
            servidorCorreo.login(correoEmisor, '$f~Pu$9zUIu)%=3')
            servidorCorreo.sendmail(correoEmisor, correoDestinatarios, mensajeCorreo.as_string())
            servidorCorreo.quit()


    ####################################
    ## Correo plataforma disfuncional ##
    ####################################


    def enviarCorreoPlataforma(self, plataforma):
        """
        Realiza el proceso del envío del correo a los interesados en caso de que una plataforma no haya funcionado.
        """

        self.tablacorreos2 = ConsultaImportante.tablaCorreoPlataforma()

        # Modificaciones iniciales a los datos de las consultas.
        self.tablaCorreos2 = pd.DataFrame(self.tablaCorreos2,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')

        # Datos sobre el correo.
        correoEmisor = 'notificaciones.sgi@appsgi.com.co'
        correoReceptor = self.tablaCorreos['correo'].dropna().tolist()
        correoCopia = self.tablaCorreos['correoCopia'].dropna().tolist()
        correoDestinatarios = correoReceptor + correoCopia
        correoAsunto = f'Notificación de errores en el ejecución de la plataforma {plataforma}'

        # Texto del correo.
        correoTexto = f"""
        Buenos días. Espero que se encuentre bien.
        
        Mediante el presente correo se le informa que la plataforma {plataforma} tuvo errores durante su ejecución. Esta ejecución se intentó varias veces sin éxito y, por ende, uno o varios archivos que esta plataforma descarga no se encuentran.
        
        Por ende, se le invita a revisar en caso de que la plataforma genuinamente presente un problema. Asimismo, el departamento de tecnología y desarrollo fue copiado en este correo y estará atento a las inquietudes o solicitudes que usted pueda tener.

        Es importante aclarar que el informe dejará todos los valores asociados a los vehículos de {plataforma} vacíos y deberá enviar los archivos derivados de esta plataforma lo más pronto posible al correo desarrollo.software@sgiltda.com.

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

    def new_method(self, plataforma):
        correoEmisor = 'notificaciones.sgi@appsgi.com.co'
        correoReceptor = self.tablaCorreos2['correo'].dropna().tolist()
        correoCopia = self.tablaCorreos2['correoCopia'].dropna().tolist()
        correoDestinatarios = correoReceptor + correoCopia
        correoAsunto = f'Notificación de errores para la plataforma {plataforma} durante el día {datetime.date.today()}'
        return correoEmisor,correoReceptor,correoCopia,correoDestinatarios,correoAsunto


    ####################################
    #### Enviar correo hora laboral ####
    ####################################


    def enviarCorreoLaboral(self):
        """
        Realiza el proceso del envío del correo al personal interesado para informarles de desplazamientos fuera de horario laboral.
        """

        self.tablaHorarios, self.tablaPuntos = ConsultaImportante().tablaCorreoLaboral()
        self.tablaExcesos, self.tablaCorreos = ConsultaImportante().tablaCorreoPersonal() # Sí, aquí se llama otra base que no se usa.

        # Modificaciones iniciales a los datos de las consultas.
        self.tablaHorarios = pd.DataFrame(self.tablaHorarios,columns=['Placa','Fecha y hora', 'Conductor'])
        self.tablaPuntos = pd.DataFrame(self.tablaHorarios,columns=['Placa', 'Conductor'])
        self.tablaPuntos = self.tablaHorarios.groupby(['Placa', 'Conductor']).size().reset_index(name='Desplazamientos fuera de hora laboral')
        self.tablaPuntos = self.tablaPuntos.sort_values(by='Desplazamientos fuera de hora laboral', ascending=False)
        self.tablaCorreos = pd.DataFrame(self.tablaCorreos,columns=['eliminar','correo','correoCopia']).drop(columns='eliminar')

        print(self.tablaPuntos)

        # Datos sobre el correo.
        correoEmisor = 'notificaciones.sgi@appsgi.com.co'
        correoReceptor = self.tablaCorreos['correo'].dropna().tolist()
        correoCopia = self.tablaCorreos['correoCopia'].dropna().tolist()
        correoDestinatarios = correoReceptor + correoCopia
        correoAsunto = f'Informe de desplazamientos fuera de horario laboral para el {datetime.date.today()}'

        # Texto del correo.
        correoTexto = f"""
        <p>Buenos d&iacute;as. Espero que se encuentre bien.</p>

        <p>Mediante el presente correo puede encontrar los desplazamientos fuera de horario que ocurrieron el {datetime.date.today()} por hora del evento.</p>

        {build_table(self.tablaHorarios, 'green_light')}

        <p>Asimismo, puede encontrar la cantidad de desplazamientos fuera de horario que tuvo cada placa.</p>

        {build_table(self.tablaPuntos, 'green_light')}

        <p>Atentamente,<br>
        Departamento de Tecnolog&iacute;a y desarrollo, SGI SAS</p>
            """
        
        # Para formatear el texto de manera correcta.
        correoTexto = textwrap.dedent(correoTexto)

        # Creación del correo.
        mensajeCorreo = MIMEMultipart()
        mensajeCorreo['From'] = f"{Header('Notificaciones SGI', 'utf-8')} <{correoEmisor}>"
        mensajeCorreo['To'] = ", ".join(correoReceptor)
        mensajeCorreo['Cc'] = ", ".join(correoCopia)
        mensajeCorreo['Subject'] = correoAsunto
        mensajeCorreo.attach(MIMEText(correoTexto, 'html'))


        # Inicializar el correo y enviar.
        servidorCorreo = smtplib.SMTP('smtp.hostinger.com', 587)
        servidorCorreo.starttls()
        servidorCorreo.login(correoEmisor, '$f~Pu$9zUIu)%=3')
        servidorCorreo.sendmail(correoEmisor, correoDestinatarios, mensajeCorreo.as_string())
        servidorCorreo.quit()



