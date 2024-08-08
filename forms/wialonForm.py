# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from db.consultasImportantes import ConsultaImportante
import pandas as pd
import os
import glob


class DatosWialon:
    def __init__(self):
        self.lugarDescargasWialon = os.getcwd() + "\\outputWialon"
  
    def rpaWialon(self):
        """
        Realiza el proceso del RPA para la plataforma Wialon.
        """

        tiempoInicio = time.time()

        self.placasPWialon = ConsultaImportante().tablaWialon()
        self.placasPWialon = pd.DataFrame(self.placasPWialon, columns=['Placa', 'plataforma'])
        self.placasWialon = self.placasPWialon['Placa'].tolist()

        # Opciones iniciales del navegador.
        opcionesNavegador = webdriver.ChromeOptions()
        if not os.path.exists(self.lugarDescargasWialon):
            os.makedirs(self.lugarDescargasWialon)

        opcionDescarga = {
            "download.default_directory": self.lugarDescargasWialon,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        }
        
        opcionesNavegador.add_argument("--headless=new")

        opcionesNavegador.add_experimental_option("prefs", opcionDescarga)
        driver = webdriver.Chrome(options= opcionesNavegador)
        driver.set_window_size(1280, 720)


        ####################################
        #### Entrada e inicio de sesión ####
        ####################################


        # Entrada a página web de Wialon
        driver.get("https://hosting.wialon.com/?lang=en")
        WebDriverWait(driver,100).until(EC.presence_of_element_located((By.ID,"LoginInputControl")))


        # Usuario
        driver.find_element(By.ID,"LoginInputControl").send_keys("DEIMER")

        # Contraseña
        driver.find_element(By.CSS_SELECTOR,".PasswordInput").send_keys("Deimer*1")

        # Botón ingreso
        driver.find_element(By.ID,"monitoringLoginMainSubmitButton").click()


        ####################################
        ##### Selección Informe General ####
        ####################################

        WebDriverWait(driver,100).until(EC.presence_of_element_located((By.XPATH,"/html/body/div[12]/div/div/div[2]/div[5]/div")))
        time.sleep(2)
        driver.find_element(By.XPATH,"/html/body/div[12]/div/div/div[2]/div[5]").click()    

        # Seleccionar template "INFORME DETALLADO POR UNIDAD".
        WebDriverWait(driver,100).until(EC.presence_of_element_located((By.ID,"report_templates_filter_reports")))
        time.sleep(2)
        driver.find_element(By.ID,"report_templates_filter_reports").click()
        time.sleep(1)
        driver.find_element(By.ID,"report_templates_filter_reports").send_keys("INFORME GENERAL")
        time.sleep(1)
        driver.find_element(By.XPATH,"/html/body/div[13]/div/div/div[3]/div/div[1]/div[1]/div[2]/div/div/div[2]/div/ul/li").click()


        ####################################
        ####### Descargar por placa ########
        ####################################

        for placa in self.placasWialon:
            # Seleccionar la placa.
            driver.find_element(By.ID,"report_templates_filter_units").click()
            time.sleep(1)
            driver.find_element(By.ID,"report_templates_filter_units").send_keys(placa)
            time.sleep(1)
            driver.find_element(By.XPATH,"/html/body/div[13]/div/div/div[3]/div/div[1]/div[2]/div[2]/div[1]/div/div/div[1]/div/div[2]/div/ul/li").click()
            time.sleep(1)

            # Oprimir el botón execute.
            WebDriverWait(driver,100).until(EC.presence_of_element_located((By.ID,"report_templates_filter_params_execute")))
            driver.find_element(By.ID,"report_templates_filter_params_execute").click()
            time.sleep(1)

            # Oprimir botón Export.
            WebDriverWait(driver,100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#report_result_export > div:nth-child(2)")))
            time.sleep(2)
            driver.find_element(By.XPATH,"/html/body/div[14]/div[6]/div/div/div[1]/div[7]/div/span/div").click()

            # Descargar en Excel para el día seleccionado. ES POSIBLE QUE SE TENGA QUE CAMBIAR PARA LOS DÍAS QUE SE PIDAN.
            WebDriverWait(driver,100).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.dropdown-option:nth-child(1)")))
            time.sleep(2)
            driver.find_element(By.CSS_SELECTOR,"div.dropdown-option:nth-child(1)").click()


        ####################################
        ####### Cierre del Webdriver #######
        ####################################


        time.sleep(1)
        archivos = glob.glob(os.path.join(self.lugarDescargasWialon, '*.xlsx'))
        self.archivoWialon1 = self.archivoWialon2 = self.archivoWialon3 = str()

        while time.time() - tiempoInicio <181: # Si se encuentran los 3 archivos en menos de 3 minutos de estar ejecutando, se acaba.
            if len(archivos) == 3:
                driver.quit()
                for archivo in archivos:
                    for placa in self.placasWialon:
                        if placa in archivo and placa == self.placasWialon[0]:
                            self.archivoWialon1 += archivo
                        else:
                            self.archivoWialon1
                        if placa in archivo and placa == self.placasWialon[1]:
                            self.archivoWialon2 += archivo
                        else:
                            self.archivoWialon2
                        if placa in archivo and placa == self.placasWialon[2]:
                            self.archivoWialon3 += archivo
                        else:
                            self.archivoWialon3
                break
            else:
                time.sleep(2)
                archivos = glob.glob(os.path.join(self.lugarDescargasWialon, '*.xlsx'))

                
        else:
            driver.quit() # Se avisa en el archivo excel para que las excepciones queden en conjunto.
        
        return self.archivoWialon1, self.archivoWialon2, self.archivoWialon3
