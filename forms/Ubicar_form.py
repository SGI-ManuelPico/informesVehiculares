# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

#class UbicarDatos:
#   def __init__(self):


opcionesNavegador = webdriver.ChromeOptions()
lugarDescargasUbicar = os.getcwd() + r"\outputUbicar"
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
####### Selección sin despls. ######
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


# Seleccionar carro
driver.find_element(By.CSS_SELECTOR,"div.row:nth-child(3) > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > button:nth-child(1)").click()
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

# Cierre del webdriver.
if len(os.listdir(lugarDescargasUbicar)) >= 2:
    driver.quit()
else:
    time.sleep(5)