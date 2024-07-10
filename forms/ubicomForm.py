# Preámbulos
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

#class ubicomDatos:
#   def __init__(self):


opcionesNavegador = webdriver.ChromeOptions()
lugarDescargasUbicom = os.getcwd() + r"\outputUbicom"
if not os.path.exists(lugarDescargasUbicom):
    os.makedirs(lugarDescargasUbicom)

opcionDescarga = {
    "download.default_directory": lugarDescargasUbicom,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}

opcionesNavegador.add_experimental_option("prefs", opcionDescarga)
driver = webdriver.Chrome(options= opcionesNavegador)
driver.set_window_size(1280, 720)


####################################
#### Entrada e inicio de sesión ####
####################################


# Entrada a página web de Ubicom
driver.get("https://gps.ubicom.co/")
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.XPATH,'//*[@id="Login"]')))


# Usuario
driver.find_element(By.XPATH,'//*[@id="Login"]').send_keys("ROLAND")

# Contraseña
driver.find_element(By.XPATH,'//*[@id="Contrasena"]').send_keys("ISV890")

# Botón ingreso
driver.find_element(By.NAME,"action").click()


####################################
### Selección sin desplazamiento ###
####################################


# Seleccionar botón informes.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.XPATH,"/html/body/ul[2]/li[2]/ul/li[2]")))
driver.find_element(By.XPATH,"/html/body/ul[2]/li[2]/ul/li[2]").click()

# Buscar el Detalle del Vehículo. Dado que es uno solo, es preferible esto al informe general.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.LINK_TEXT,"Detalle Vehículo")))
driver.find_element(By.LINK_TEXT,"Detalle Vehículo").click()

# Buscar el Vehículo de esta plataforma. Por razones de pruebas, se usará FNM236, pero será necesario conectarlo a una base de datos después.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID,"ddlVehiculo")))
driver.find_element(By.ID,"ddlVehiculo").click()
driver.find_element(By.XPATH,"/html/body/main/div[1]/div[1]/form/div[1]/select/option[12]").click()


####################################
###### Descargar información #######
####################################


# Consultar el detalle del vehículo en el día actual. ES POSIBLE QUE SE TENGA QUE CAMBIAR PARA LOS DÍAS QUE SE PIDAN.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID, "btnConsultar")))
driver.find_element(By.ID, "btnConsultar").click()

# Descargar el detalle del vehículo en el día actual. ES POSIBLE QUE SE TENGA QUE CAMBIAR PARA LOS DÍAS QUE SE PIDAN.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID,"btnExportarEXCEL")))
driver.find_element(By.ID,"btnExportarEXCEL").click()


####################################
### Selección de desplazamientos ###
####################################


# Seleccionar botón informes.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.XPATH,"/html/body/ul[2]/li[2]/ul/li[2]")))
driver.find_element(By.XPATH,"/html/body/ul[2]/li[2]/ul/li[2]").click()

# Buscar la información de "Estacionados".
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.LINK_TEXT,"Estacionados")))
driver.find_element(By.LINK_TEXT,"Estacionados").click()

# Buscar el Vehículo de esta plataforma. Por razones de pruebas, se usará FNM236, pero será necesario conectarlo a una base de datos después.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID,"ddlVehiculo")))
driver.find_element(By.ID,"ddlVehiculo").click()
driver.find_element(By.XPATH,"/html/body/main/div[1]/div[1]/form/div[1]/select/option[12]").click()


####################################
#### Descargar despls. y cerrar ####
####################################


# Consultar el detalle del vehículo en el día actual. ES POSIBLE QUE SE TENGA QUE CAMBIAR PARA LOS DÍAS QUE SE PIDAN.
driver.find_element(By.ID, "btnConsultar").click()

# Descargar el detalle del vehículo en el día actual. ES POSIBLE QUE SE TENGA QUE CAMBIAR PARA LOS DÍAS QUE SE PIDAN.
WebDriverWait(driver,50).until(EC.presence_of_element_located((By.ID,"btnExportarEXCEL")))
driver.find_element(By.ID,"btnExportarEXCEL").click()

# Cierre del webdriver.
archivoUbicom1 = os.getcwd() + "\\outputUbicom\\ReporteDiario.xls"
archivoUbicom2 = os.getcwd() + "\\outputUbicom\\Estacionados.xls"
while os.path.isfile(archivoUbicom1) and os.path.isfile(archivoUbicom2):
    time.sleep(2)
    driver.quit()
else:
    time.sleep(2)