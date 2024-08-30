import os
import shutil
import win32com.client as win32
from datetime import datetime

class TratadorArchivos:
    def __init__(self):
        pass


    def eliminarArchivosOutput(self):
        """
        Elimina los archivos que aparecen en las carpetas de Output de cada RPA de cada plataforma.
        """
        for folder in os.listdir():
            if "output" in folder:
                shutil.rmtree(folder)


    def eliminarArchivosPlataforma(self, plataforma):
        """
        Elimina los archivos que aparecen en la carpeta de Output de una plataforma.
        """
        carpetaOutput = os.getcwd() + "\\output" + plataforma
        print(carpetaOutput)
        shutil.rmtree(carpetaOutput)

    
    def xlsx(self, input_file):
        """
        Hace que un archivo se vuelva formato xlsx.
        """
        # Start an instance of Excel
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        
        # Open the XLS file
        wb = excel.Workbooks.Open(r"{}".format(input_file))
        
        # Save the file as XLSX
        self.temp_file = os.path.splitext(input_file)[0] + '.xlsx'
        wb.SaveAs(os.path.abspath(self.temp_file), FileFormat=51)  # FileFormat=51 is for .xlsx
        wb.Close()
        excel.Quit()
        return self.temp_file

    def conversorSegundosWialon(duration_str):
        parts = duration_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        else:
            return int(parts[0])


    def conversorSegundosUbicar(duration):
        parts = duration.split(' ')
        total_seconds = 0
        for part in parts:
            if 'h' in part:
                total_seconds += int(part.replace('h', '')) * 3600
            elif 'min' in part:
                total_seconds += int(part.replace('min', '')) * 60
            elif 's' in part:
                total_seconds += int(part.replace('s', ''))
        return total_seconds


    def conversorSegundosMDVR(duration_str):
        parts = duration_str.split(' ')
        minutes = 0
        seconds = 0
        for part in parts:
            if 'min' in part:
                minutes += int(part.replace('min', ''))
            if 's' in part:
                seconds += int(part.replace('s', ''))
        return minutes * 60 + seconds
    
    def crearDirectorioError(self, plataforma):
        """
        Crea un directorio con el nombre de la plataforma, y un subdirectorio dentro de este on la fecha en la que ocurre el error.
        """
        current_date = datetime.now().strftime("%d-%m")
        directorio_plataforma = plataforma
        directorio_fecha = f"{directorio_plataforma}/{current_date}"
        
        if not os.path.exists(directorio_plataforma):
            os.makedirs(directorio_plataforma)
            os.makedirs(directorio_fecha)
        else:
            if not os.path.exists(directorio_fecha):
                os.makedirs(directorio_fecha)
            else:
                print(f"Ya existe el directorio {directorio_fecha}.")

