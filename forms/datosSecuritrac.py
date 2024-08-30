# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os

class DatosSecuritrac:
    def __init__(self):
        self.archivoSecuritrac = os.getcwd() + "\\outputSecuritrac\\exported-excel.xls"
        self.lugarDescargasSecuritrac = os. getcwd() + r"\outputSecuritrac"


    def rpaSecuritrac(self):
        """
        Realiza el proceso del RPA para la plataforma Securitrac.
        """

        tiempoInicio = time.time()

        opcionesNavegador = webdriver.ChromeOptions()
        if not os.path.exists(self.lugarDescargasSecuritrac):
            os.makedirs(self.lugarDescargasSecuritrac)

        opcionDescarga = {
            "download.default_directory": self.lugarDescargasSecuritrac,
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


        # Entrada a página web de Secturitrac
        driver.get("https://www.securitrac.net/web/#!login")
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.ID,"gwt-uid-3")))


        # Usuario
        driver.find_element(By.ID,"gwt-uid-3").send_keys("SGIGAITAN")

        # Contraseña
        driver.find_element(By.ID,"gwt-uid-5").send_keys("SGIGAITAN")

        # Botón ingreso
        driver.find_element(By.ID,"gwt-uid-5").send_keys(Keys.ENTER)


        ####################################
        ######## Buscar información ########
        ####################################


        # Seleccionar botón mora xd.

        #WebDriverWait(driver,50).until(EC.presence_of_element_located((By.XPATH,"/html/body/div[2]/div[2]/div/div/div[3]/div/div/div/div[3]/div"))):
        #time.sleep(2)    
        #driver.find_element(By.XPATH,"/html/body/div[2]/div[2]/div/div/div[3]/div/div/div/div[3]/div").click()

        # Seleccionar todos los vehículos
        
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.XPATH,"/html/body/div[1]/div/div[2]/div/div/div[3]/div/div/div[3]/div/div/div[2]/div/div[3]/table/thead/tr/th[1]/span/input")))
        time.sleep(1)
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div[3]/div/div/div[3]/div/div/div[2]/div/div[3]/table/thead/tr/th[1]/span/input").click()

        # Seleccionar botón Informes Eventos.
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/div/div/div[1]/div/div/div[2]/div/span[2]").click()
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.CSS_SELECTOR,".v-menubar-submenu > span:nth-child(3)")))
        driver.find_element(By.CSS_SELECTOR,".v-menubar-submenu > span:nth-child(3)").click()

        # Seleccionar Eventos
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.CSS_SELECTOR,"span.v-checkbox:nth-child(2) > label:nth-child(2)")))
        time.sleep(2)
        driver.find_element(By.CSS_SELECTOR,"span.v-checkbox:nth-child(2)").click()
        driver.find_element(By.CSS_SELECTOR,"span.v-checkbox:nth-child(4)").click()
        driver.find_element(By.CSS_SELECTOR,"span.v-checkbox:nth-child(5)").click()
        driver.find_element(By.CSS_SELECTOR,"div.v-gridlayout-slot:nth-child(5) > div:nth-child(1)").click()


        ####################################
        ######## Descargar y Cerrar ########
        ####################################


        # Dado que son todos los vehículos de esta plataforma, se dejará así por motivos de pruebas, pero tocará cambiar esto.
        WebDriverWait(driver,50).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.v-slot:nth-child(5) > div:nth-child(1)"))) # Verifica si el botón único de "Exportar a KML está presente."
        driver.find_element(By.XPATH,"/html/body/div[1]/div/div[2]/div/div/div[3]/div/div/div[1]/div/div/div[1]/div/div[2]/div/div/div/div/div[2]/div/div[1]/div").click()

        # Cierre del webdriver.
        while time.time() - tiempoInicio <181:
            if os.path.isfile(self.archivoSecuritrac) == True:
                time.sleep(2)
                driver.quit()
                break
            else:
                time.sleep(2)
        else:
            driver.quit() # Se avisa en el archivo excel para que las excepciones queden en conjunto.
        
        return self.archivoSecuritrac