# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import os
import glob

# class MDVRDatos:
#     def __init__(self):
#         lugarDescargasMDVR = os.getcwd() + r"\outputMDVR"
#         archivos = glob.glob(os.path.join(lugarDescargasMDVR, '*.xls'))

lugarDescargasMDVR = os.getcwd() + r"\outputMDVR"
archivos = glob.glob(os.path.join(lugarDescargasMDVR, '*.xls'))


def rpaMDVR():
    """
    Realiza el proceso del RPA para la plataforma MDVR.
    """
    opcionesNavegador = webdriver.ChromeOptions()
    if not os.path.exists(lugarDescargasMDVR):
        os.makedirs(lugarDescargasMDVR)

    opcionDescarga = {
        "download.default_directory": lugarDescargasMDVR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    }

    opcionesNavegador.add_experimental_option("prefs", opcionDescarga)
    driver = webdriver.Chrome(options= opcionesNavegador)
    driver.set_window_size(1280, 720)


    ####################################
    #### Entrada e inicio de sesión ####
    ####################################


    # Entrada a página web de MDVR
    driver.get("https://mdvrgps.ddns.net/authentication/create")
    WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID,"sign-in-form-email")))


    # Usuario
    driver.find_element(By.ID,"sign-in-form-email").send_keys("johana298@mdvr.com")

    # Contraseña
    driver.find_element(By.ID,"sign-in-form-password").send_keys("ksz298")

    # Botón ingreso
    driver.find_element(By.NAME,"Submit").click()


    ####################################
    ####### Selección sin despls. ######
    ####################################


    # Seleccionar Herramientas
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.ID,"dropTools")))
    time.sleep(2)
    driver.find_element(By.ID,"dropTools").click()

    # Seleccionar Reportes
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[2]/nav/div/div[2]/ul/li[1]/ul/li[4]/a")))
    driver.find_element(By.XPATH,"/html/body/div[2]/nav/div/div[2]/ul/li[1]/ul/li[4]/a").click()

    # Seleccionar Formato correcto (XLSX)
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,".generate")))
    driver.find_element(By.CSS_SELECTOR,"div.col-sm-6:nth-child(3) > div:nth-child(1) > div:nth-child(2) > button:nth-child(1)").click()
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"div.col-sm-6:nth-child(3) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(2)")))
    driver.find_element(By.CSS_SELECTOR,"div.col-sm-6:nth-child(3) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(2)").click()

    # Introducir límite de velocidad
    driver.find_element(By.NAME, "speed_limit").send_keys(80)

    # Seleccionar carro
    driver.find_element(By.XPATH,"/html/body/div[12]/div/div/div/div/div[2]/div/form/div/div[1]/div[2]/div[2]/div[1]/div/div/button/span[2]").click()
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"div.open:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > button:nth-child(1)")))
    driver.find_element(By.CSS_SELECTOR,"div.open:nth-child(1) > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > button:nth-child(1)").click()

    # Descargar
    driver.find_element(By.CSS_SELECTOR,".generate").click()
    time.sleep(1)


    ####################################
    #### Selección despls. y cerrar ####
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
    WebDriverWait(driver,50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#reports-form-reports > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(2) > ul:nth-child(1) > li:nth-child(17)")))
    driver.find_element(By.XPATH,"/html/body/div[12]/div/div/div/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/div/div/ul/li[16]/a").send_keys(Keys.PAGE_DOWN)
    time.sleep(1)
    driver.find_element(By.XPATH,"/html/body/div[12]/div/div/div/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/div/div/ul/li[17]/a").click()

    # Descargar
    driver.find_element(By.CSS_SELECTOR,".generate").click()
    time.sleep(1)

    # Cierre del webdriver.
    while len(archivos) == 3:
        driver.quit()
    else:
        time.sleep(2)

archivoMDVR1 = str()
archivoMDVR2 = str()
archivoMDVR3 = str()
for archivo in archivos:
    if "general" in archivo:
        archivoMDVR1 += archivo
    elif "drivers" in archivo:
        archivoMDVR2 += archivo
    else:
        archivoMDVR3 += archivo
