from db.consultasImportantes import ConsultaImportante
from forms.ituranForm import DatosIturan
from forms.MDVRForm import DatosMDVR
from forms.securitracForm import DatosSecuritrac
from forms.ubicarForm import DatosUbicar
from forms.ubicomForm import DatosUbicom
from forms.wialonForm import DatosWialon
from util.correosVehiculares import CorreosVehiculares
########################
import pandas as pd



class RPA:
    def __init__(self):
        pass

    def ejecutarRPAIturan(self):
        """
        Define la secuencia de ejecución del RPA de Ituran.
        """

        ConsultaImportante.tablaEstadosPlataforma("Ituran")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=['plataforma', "estado"])

        if self.tablaEstados.iloc[0]["estado"]=="Ejecutado":
            pass
        else:
            try:
                self.archivoIturan1, self.archivoIturan2, self.archivoIturan3 = DatosIturan.rpaIturan()
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ituran.")
                self.tablaEstados.at[0]['estado'] = "Error"
                # CorreosVehiculares.enviarCorreoPlataforma("Ituran")

    def ejecutarRPASecuritrac(self):
        """
        Define la secuencia de ejecución del RPA de Securitrac.
        """

        ConsultaImportante.tablaEstadosPlataforma("Securitrac")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=['plataforma', "estado"])

        if self.tablaEstados.iloc[0]["estado"]=="Ejecutado":
            pass
        else:
            try:
                self.archivoSecuritrac = DatosSecuritrac.rpaSecuritrac()
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ituran.")
                self.tablaEstados.at[0]['estado'] = "Error"
                # CorreosVehiculares.enviarCorreoPlataforma("Securitrac")

    def ejecutarRPAMDVR(self):
        """
        Define la secuencia de ejecución del RPA de MDVR.
        """

        ConsultaImportante.tablaEstadosPlataforma("MDVR")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=['plataforma', "estado"])

        if self.tablaEstados.iloc[0]["estado"]=="Ejecutado":
            pass
        else:
            try:
                self.archivoMDVR1, self.archivoMDVR2, self.archivoMDVR3 = DatosMDVR.rpaMDVR()
            except:
                print("Hubo un error en el acceso por el internet para ingresar a MDVR.")
                self.tablaEstados.at[0]['estado'] = "Error"
                # CorreosVehiculares.enviarCorreoPlataforma("MDVR")

    def ejecutarRPAUbicar(self):
        """
        Define la secuencia de ejecución del RPA de Ubicar.
        """

        ConsultaImportante.tablaEstadosPlataforma("Ubicar")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=['plataforma', "estado"])

        if self.tablaEstados.iloc[0]["estado"]=="Ejecutado":
            pass
        else:
            try:
                self.archivoUbicar1, self.archivoUbicar2, self.archivoUbicar3 = DatosUbicar.rpaUbicar()
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ubicar.")
                self.tablaEstados.at[0]['estado'] = "Error"
                # CorreosVehiculares.enviarCorreoPlataforma("Ubicar")

    def ejecutarRPAUbicom(self):
        """
        Define la secuencia de ejecución del RPA de Ubicom.
        """

        ConsultaImportante.tablaEstadosPlataforma("Ubicom")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=['plataforma', "estado"])

        if self.tablaEstados.iloc[0]["estado"]=="Ejecutado":
            pass
        else:
            try:
                self.archivoUbicom1, self.archivoUbicom2 = DatosUbicom.rpaUbicom()
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ubicom.")
                self.tablaEstados.at[0]['estado'] = "Error"
                # CorreosVehiculares.enviarCorreoPlataforma("Ubicom")

    def ejecutarRPAWialon(self):
        """
        Define la secuencia de ejecución del RPA de Wialon.
        """

        ConsultaImportante.tablaEstadosPlataforma("Wialon")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=['plataforma', "estado"])

        if self.tablaEstados.iloc[0]["estado"]=="Ejecutado":
            pass
        else:
            try:
                self.archivoWialon1, self.archivoWialon2, self.archivoWialon3 = DatosWialon.rpaWialon()
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Wialon.")
                self.tablaEstados.at[0]['estado'] = "Error"
                # CorreosVehiculares.enviarCorreoPlataforma("Wialon")

