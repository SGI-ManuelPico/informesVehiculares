# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import glob

# class ubicarDatos:
#     def __init__(self):
#         lugarDescargasUbicar = os.getcwd() + r"\outputUbicar"
#         archivos = glob.glob(os.path.join(lugarDescargasUbicar, '*.xlsx'))

lugarDescargasUbicar = os.getcwd() + r"\outputUbicar"
archivos = glob.glob(os.path.join(lugarDescargasUbicar, '*.xlsx'))

def rpaUbicar():
    """
    Realiza el proceso del RPA para la plataforma Ubicar.
    """
    opcionesNavegador = webdriver.ChromeOptions()
    if not os.path.exists(lugarDescargasUbicar):
        os.makedirs(lugarDescargasUbicar)

    opcionDescarga = {
        "download.default_directory": lugarDescargasUbicar,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    }

    opcionesNavegador.add_experimental_option("prefs", opcionDescarga)
    driver = webdriver.Chrome(options= opcionesNavegador)
    driver.set_window_size(1280, 720)


    ####################################
    #### Entrada e inicio de sesión ####
    ####################################


    # Entrada a página web de Ubicar
    driver.get("https://plataforma.sistemagps.online/authentication/create")
    WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID,"sign-in-form-email")))


    # Usuario
    driver.find_element(By.ID,"sign-in-form-email").send_keys("jyt620@ubicar.gps")

    # Contraseña
    driver.find_element(By.ID,"sign-in-form-password").send_keys("123456")

    # Botón ingreso
    driver.find_element(By.XPATH,"/html/body/div/div/div/div/div[2]/form/button").click()


    ####################################
    ###### Selección excesos vel. ######
    ####################################


    # Seleccionar Herramientas
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.ID,"dropTools")))
    driver.find_element(By.ID,"dropTools").click()

    # Seleccionar Reportes
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[2]/nav/div/div[2]/ul/li[1]/ul/li[4]/a")))
    driver.find_element(By.XPATH,"/html/body/div[2]/nav/div/div[2]/ul/li[1]/ul/li[4]/a").click()

    # Seleccionar Formato correcto (XLSX)
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,".generate")))
    driver.find_element(By.CSS_SELECTOR,"div.col-sm-6:nth-child(3) > div:nth-child(1) > div:nth-child(2) > button:nth-child(1)").click()
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"div.col-sm-6:nth-child(3) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4)")))
    driver.find_element(By.CSS_SELECTOR,"div.col-sm-6:nth-child(3) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(4)").click()

    # Introducir límite de velocidad
    driver.find_element(By.NAME, "speed_limit").send_keys(80)

    # Seleccionar carro
    driver.find_element(By.CSS_SELECTOR,"div.row:nth-child(3) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > button:nth-child(1)").click()
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"div.open:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > button:nth-child(1)")))
    driver.find_element(By.CSS_SELECTOR,"div.open:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > button:nth-child(1)").click()

    # Descargar
    driver.find_element(By.CSS_SELECTOR,".generate").click()
    time.sleep(1)


    ####################################
    # Selección despls. y Kms y cerrar #
    ####################################


    # Seleccionar Tipo y "Recorridos y paradas"
    driver.find_element(By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > button:nth-child(1)").click()
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(9) > a:nth-child(1)")))
    driver.find_element(By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(9) > a:nth-child(1)").click()

    # Descargar
    driver.find_element(By.CSS_SELECTOR,".generate").click()
    time.sleep(1)


    ####################################
    #### Selección excesos y cerrar ####
    ####################################


    # Seleccionar Tipo y "Excesos de velocidad"
    driver.find_element(By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > button:nth-child(1)").click()
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div > div > div > ul > li:nth-child(20) > a")))
    driver.find_element(By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div > div > div > ul > li:nth-child(20) > a").click()
    time.sleep(1)

    # Descargar
    driver.find_element(By.CSS_SELECTOR,".generate").click()
    time.sleep(1)

    # Cierre del webdriver.
    if len(archivos) == 3:
        time.sleep(5)
        driver.quit()
    else:
        time.sleep(2)

archivoUbicar1 = str()
archivoUbicar2 = str()
archivoUbicar3 = str()
if len(archivos) == 3:
    for archivo in archivos:
        if "general" in archivo:
            archivoUbicar1 += archivo
        elif "stops" in archivo:
            archivoUbicar2 += archivo
        else:
            archivoUbicar3 += archivo
