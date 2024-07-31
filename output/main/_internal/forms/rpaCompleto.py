from db.consultasImportantes import ConsultaImportante
from forms.ituranForm import DatosIturan
from forms.MDVRForm import DatosMDVR
from forms.securitracForm import DatosSecuritrac
from forms.ubicarForm import DatosUbicar
from forms.ubicomForm import DatosUbicom
from forms.wialonForm import DatosWialon
from util.tratadoArchivos import TratadorArchivos
import os, glob
import pandas as pd



class RPA:
    def __init__(self):
        pass


    def ejecutarRPAIturan(self):
        """
        Define la secuencia de ejecución del RPA de Ituran.
        """

        self.tablaEstados = ConsultaImportante().tablaEstadosPlataforma(plataforma="Ituran")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=["estado"])

        if self.tablaEstados.iloc[0]["estado"]!="Ejecutado":
            try:
                self.archivoIturan1, self.archivoIturan2, self.archivoIturan3, self.archivoIturan4 = DatosIturan().rpaIturan()
                ConsultaImportante().actualizarEstadoPlataforma("Ituran","Ejecutado")
                return self.archivoIturan1, self.archivoIturan2, self.archivoIturan3, self.archivoIturan4
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ituran.")
                TratadorArchivos().elimnarArchivosPlataforma("Ituran")
                ConsultaImportante().actualizarEstadoPlataforma("Ituran","Error")
                self.archivoIturan1 = self.archivoIturan2 = self.archivoIturan3 = self.archivoIturan4 = os.getcwd() + r"\archivoFicticio.csv"
                return self.archivoIturan1, self.archivoIturan2, self.archivoIturan3, self.archivoIturan4


    def ejecutarRPASecuritrac(self):
        """
        Define la secuencia de ejecución del RPA de Securitrac.
        """

        self.tablaEstados = ConsultaImportante().tablaEstadosPlataforma(plataforma="Securitrac")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=["estado"])

        if self.tablaEstados.iloc[0]["estado"]!="Ejecutado":
            try:
                self.archivoSecuritrac = DatosSecuritrac().rpaSecuritrac()
                ConsultaImportante().actualizarEstadoPlataforma("Securitrac","Ejecutado")
                return self.archivoSecuritrac
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Securitrac.")
                TratadorArchivos().eliminarArchivosPlataforma("Securitrac")
                ConsultaImportante().actualizarEstadoPlataforma("Securitrac","Error")
                self.archivoSecuritrac =os.getcwd() + r"\archivoFicticio.xls"
                return self.archivoSecuritrac


    def ejecutarRPAMDVR(self):
        """
        Define la secuencia de ejecución del RPA de MDVR.
        """

        self.tablaEstados = ConsultaImportante().tablaEstadosPlataforma(plataforma="MDVR")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=["estado"])

        if self.tablaEstados.iloc[0]["estado"]!="Ejecutado":
            try:
                self.archivoMDVR1, self.archivoMDVR2, self.archivoMDVR3 = DatosMDVR().rpaMDVR()
                ConsultaImportante().actualizarEstadoPlataforma("MDVR","Ejecutado")
                return self.archivoMDVR1, self.archivoMDVR2, self.archivoMDVR3
            except:
                print("Hubo un error en el acceso por el internet para ingresar a MDVR.")
                TratadorArchivos().eliminarArchivosPlataforma("MDVR")
                ConsultaImportante().actualizarEstadoPlataforma("MDVR","Error")
                self.archivoMDVR1 = self.archivoMDVR2 = self.archivoMDVR3= os.getcwd() + r"\archivoFicticio.xlsx"
                return self.archivoMDVR1, self.archivoMDVR2, self.archivoMDVR3


    def ejecutarRPAUbicar(self):
        """
        Define la secuencia de ejecución del RPA de Ubicar.
        """

        self.tablaEstados = ConsultaImportante().tablaEstadosPlataforma(plataforma="Ubicar")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=["estado"])

        if self.tablaEstados.iloc[0]["estado"]!="Ejecutado":
            try:
                self.archivoUbicar1, self.archivoUbicar2, self.archivoUbicar3 = DatosUbicar().rpaUbicar()
                ConsultaImportante().actualizarEstadoPlataforma("Ubicar","Ejecutado")
                return self.archivoUbicar1, self.archivoUbicar2, self.archivoUbicar3
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ubicar.")
                TratadorArchivos().eliminarArchivosPlataforma("Ubicar")
                ConsultaImportante().actualizarEstadoPlataforma("Ubicar","Error")
                self.archivoUbicar1 = self.archivoUbicar2 = self.archivoUbicar3= os.getcwd() + r"\archivoFicticio.xlsx"
                return self.archivoUbicar1, self.archivoUbicar2, self.archivoUbicar3


    def ejecutarRPAUbicom(self):
        """
        Define la secuencia de ejecución del RPA de Ubicom.
        """

        self.tablaEstados = ConsultaImportante().tablaEstadosPlataforma(plataforma="Ubicom")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=["estado"])

        if self.tablaEstados.iloc[0]["estado"]!="Ejecutado":
            try:
                self.archivoUbicom1, self.archivoUbicom2 = DatosUbicom().rpaUbicom()
                ConsultaImportante().actualizarEstadoPlataforma("Ubicom","Ejecutado")
                return self.archivoUbicom1, self.archivoUbicom2
            except:
                print("Hubo un error en el acceso por el internet para ingresar a Ubicom.")
                TratadorArchivos().eliminarArchivosPlataforma("Ubicom")
                ConsultaImportante().actualizarEstadoPlataforma("Ubicom","Error")
                self.archivoUbicom1 = self.archivoUbicom2 = os.getcwd() + r"\archivoFicticio.xls"
                return self.archivoUbicom1, self.archivoUbicom2


    def ejecutarRPAWialon(self):
        """
        Define la secuencia de ejecución del RPA de Wialon.
        """

        self.tablaEstados = ConsultaImportante().tablaEstadosPlataforma(plataforma="Wialon")

        self.tablaEstados = pd.DataFrame(self.tablaEstados, columns=["estado"])

        if self.tablaEstados.iloc[0]["estado"]!="Ejecutado":
            try:
                self.archivoWialon1, self.archivoWialon2, self.archivoWialon3 = DatosWialon().rpaWialon()
                ConsultaImportante().actualizarEstadoPlataforma("Wialon","Ejecutado")
                return self.archivoWialon1, self.archivoWialon2, self.archivoWialon3
            except Exception as e:
                print(e)
                print("Hubo un error en el acceso por el internet para ingresar a Wialon.")
                TratadorArchivos().eliminarArchivosPlataforma("Wialon")
                ConsultaImportante().actualizarEstadoPlataforma("Wialon","Error")
                self.archivoWialon1 = self.archivoWialon2 = self.archivoWialon3= os.getcwd() + r"\archivoFicticio.xlsx"
                return self.archivoWialon1, self.archivoWialon2, self.archivoWialon3