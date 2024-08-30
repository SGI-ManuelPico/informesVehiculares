# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

class DatosIturan():
    def __init__(self):
        self.archivoIturan1 = os.getcwd() + "\\outputIturan\\report.csv"
        self.archivoIturan2 = os.getcwd() + "\\outputIturan\\report (1).csv"
        self.archivoIturan3 = os.getcwd() + "\\outputIturan\\report (2).csv"
        self.archivoIturan4 = os.getcwd() + "\\outputIturan\\report (3).csv"
        self.lugarDescargasIturan = os.getcwd() + "\\outputIturan"


    def rpaIturan(self):
        """
        Realiza el proceso del RPA para la plataforma Ituran.
        """

        tiempoInicio = time.time()


        ####################################
        #### Entrada e inicio de sesión ####
        ####################################


        opcionesNavegador = webdriver.ChromeOptions()
        if not os.path.exists(self.lugarDescargasIturan):
            os.makedirs(self.lugarDescargasIturan)

        opcionDescarga = {
            "download.default_directory": self.lugarDescargasIturan,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
        }

        opcionesNavegador.add_argument("--headless=new")

        # Opciones del navegador
        opcionesNavegador.add_experimental_option("prefs", opcionDescarga)
        driver = webdriver.Chrome(options= opcionesNavegador)
        driver.set_window_size(1280, 720)

        # Entrada a página web de Ituran
        driver.get("https://www.worldfleetlog.com/ituran/")
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.ID,"UserNameText")))


        # Usuario
        driver.find_element(By.ID,"UserNameText").send_keys("SGIGPS@SGI")

        # Contraseña
        driver.find_element(By.ID,"PasswordText").send_keys("Gps987@!")

        # Botón ingreso
        driver.find_element(By.ID,"btnLogon").click()


        ####################################
        ##### Selección Informe General ####
        ####################################


        # Seleccionar "Reportes de flota".
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.boxSize:nth-child(5) > div:nth-child(1) > div:nth-child(1)")))
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR,"div.boxSize:nth-child(5) > div:nth-child(1) > div:nth-child(1)").click()
        time.sleep(1)

        # Seleccionar "Exceso de velocidad por vehículo (resumen)"
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.CSS_SELECTOR,".rpRootGroup > li:nth-child(1) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(1) > a:nth-child(1)")))
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR,".rpRootGroup > li:nth-child(1) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(1) > a:nth-child(1)").click()
        time.sleep(1)


        ####################################
        ###### Descargar información #######
        ####################################


        # Oprimir botón exportar.
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.NAME, "RadWndCreteria")))
        driver.switch_to.frame(driver.find_element(By.NAME, "RadWndCreteria"))
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.ID, "imgExportArrow")))
        time.sleep(1)
        driver.find_element(By.ID,"imgExportArrow").click()

        # Descargar en formato CSV.
        WebDriverWait(driver,500).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)")))
        driver.find_element(By.CSS_SELECTOR,".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)").click()
        time.sleep(5)

        # Verificar si puede salirse de la descarga.
        while time.time() - tiempoInicio <91:
            if os.path.isfile(self.archivoIturan1) == True:
                break
            else:
                time.sleep(2)
        else:
            driver.quit() # Se avisa en el archivo excel para que las excepciones queden en conjunto.

        ####################################
        #### Descargar Informe Detallado ###
        ####################################


        driver.switch_to.default_content()

        # Seleccionar "Exceso de velocidad por vehículo (detallado)"
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.XPATH,"/html/body/form/div[4]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div/ul/li[1]/div/ul/li[2]")))
        time.sleep(1)
        driver.find_element(By.XPATH,"/html/body/form/div[4]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div/ul/li[1]/div/ul/li[2]").click()
        time.sleep(1)

        # Oprimir botón exportar.
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.NAME, "RadWndCreteria")))
        driver.switch_to.frame(driver.find_element(By.NAME, "RadWndCreteria"))
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.ID, "imgExportArrow")))
        time.sleep(1)
        driver.find_element(By.ID,"imgExportArrow").click()

        # Descargar en formato CSV.
        WebDriverWait(driver,500).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)")))
        driver.find_element(By.CSS_SELECTOR,".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)").click()
        time.sleep(5)

        # Verificar si puede salirse de la descarga.
        while time.time() - tiempoInicio <136:
            if os.path.isfile(self.archivoIturan2) == True:
                break
            else:
                time.sleep(2)
        else:
            driver.quit() # Se avisa en el archivo excel para que las excepciones queden en conjunto.

        ####################################
        #### Descargar Distancia Diaria ####
        ####################################


        driver.switch_to.default_content()

        # Seleccionar "Distancia diaria de vehículos"
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.XPATH,"/html/body/form/div[4]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div/ul/li[1]/div/ul/li[7]")))
        time.sleep(1)
        driver.find_element(By.XPATH,"/html/body/form/div[4]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div/ul/li[1]/div/ul/li[7]").click()
        time.sleep(1)

        # Oprimir botón exportar.
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.NAME, "RadWndCreteria")))
        driver.switch_to.frame(driver.find_element(By.NAME, "RadWndCreteria"))
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.ID, "imgExportArrow")))
        time.sleep(1)
        driver.find_element(By.ID,"imgExportArrow").click()

        # Descargar en formato CSV.
        WebDriverWait(driver,500).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)")))
        driver.find_element(By.CSS_SELECTOR,".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)").click()
        time.sleep(5)

        while time.time() - tiempoInicio <181:
            if os.path.isfile(self.archivoIturan3) == True:
                time.sleep(2)
                break
            else:
                time.sleep(2)
        else:
            driver.quit() # Se avisa en el archivo excel para que las excepciones queden en conjunto.

        ####################################
        ### Descargar Viajes de vehículo ###
        ####################################


        driver.switch_to.default_content()

        # Seleccionar ""
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.XPATH,"/html/body/form/div[4]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div/ul/li[1]/div/ul/li[3]")))
        time.sleep(1)
        driver.find_element(By.XPATH,"/html/body/form/div[4]/div/div/div/table/tbody/tr/td[1]/div/table/tbody/tr[2]/td/div/ul/li[1]/div/ul/li[3]").click()
        time.sleep(1)

        # Oprimir botón exportar.
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.NAME, "RadWndCreteria")))
        driver.switch_to.frame(driver.find_element(By.NAME, "RadWndCreteria"))
        WebDriverWait(driver,500).until(EC.presence_of_element_located((By.ID, "imgExportArrow")))
        time.sleep(1)
        driver.find_element(By.ID,"imgExportArrow").click()

        # Descargar en formato CSV.
        WebDriverWait(driver,500).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)")))
        driver.find_element(By.CSS_SELECTOR,".slidingDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1)").click()
        time.sleep(2)

        ####################################
        ####### Cierre del Webdriver #######
        ####################################


        while time.time() - tiempoInicio <226:
            if os.path.isfile(self.archivoIturan4) == True:
                time.sleep(2)
                driver.quit()
                break
            else:
                time.sleep(2)
        else:
            driver.quit() # Se avisa en el archivo excel para que las excepciones queden en conjunto.

        return self.archivoIturan1, self.archivoIturan2, self.archivoIturan3, self.archivoIturan4
