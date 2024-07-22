import os
import shutil


class EliminarArchivos:
    def __init__(self):
        pass

    ####################################
    ##### Eliminar archivos del día ####
    ####################################


    def eliminarArchivosOutput(self):
        """
        Elimina los archivos que aparecen en las carpetas de Output de cada RPA de cada plataforma.
        """
        for folder in os.listdir():
            if "output" in folder:
                shutil.rmtree(folder)


    ####################################
    ##### Eliminar archivos carpeta ####
    ####################################


    def eliminarArchivosPlataforma(self, plataforma):
        """
        Elimina los archivos que aparecen en la carpeta de Output de una plataforma.
        """
        carpetaOutput = os.getcwd() + "\\output" + plataforma
        shutil.rmtree(carpetaOutput)

