import pandas as pd
import os
#from openpyxl import workbook, worksheet
import win32com.client as win32

####################################
#### Datos de placas Securitrac ####
####################################


# Desplazamientos
securitracDesplazamientos = pd.read_excel(r"outputSecturitrac\exported-excel.xls")
securitracDesplazamientos = securitracDesplazamientos[['NROMOVIL', 'EVENTO']].loc[securitracDesplazamientos['EVENTO'] == "Apagado"].value_counts().to_frame()
securitracDesplazamientos = securitracDesplazamientos.reset_index().rename(columns={"NROMOVIL": "Placa", "count": "Desplazamientos"}).drop('EVENTO', axis=1)

# Kilometraje
securitracKilometraje = pd.read_excel(r"outputSecturitrac\exported-excel.xls", usecols='A,H')
securitracKilometraje = securitracKilometraje.groupby(['NROMOVIL']).max().reset_index()
securitracKilometraje = securitracKilometraje.rename(columns={"NROMOVIL":"Placa", "KILOMETROS":"Kilometraje"})

# Excesos de velocidad
securitracExcesos = pd.read_excel(r"outputSecturitrac\exported-excel.xls")
securitracExcesos = securitracExcesos[['NROMOVIL', 'EVENTO']].loc[securitracExcesos['EVENTO'] == "Exc. Velocidad"].value_counts().to_frame()
securitracExcesos = securitracExcesos.reset_index().rename(columns={"NROMOVIL": "Placa", "count": "Excesos de velocidad"}).drop('EVENTO', axis=1)

securitracCompleto = pd.merge(securitracExcesos, securitracKilometraje, on='Placa', how='outer').merge(securitracDesplazamientos, on='Placa', how='outer').fillna(0)


####################################
###### Datos de placas Wialon ######
####################################


placasWialon = ["JTV645", "LPN816", "LPN821"]

wialonCompleto = pd.DataFrame()
lugarDescargasWialon = os.getcwd() + "\\outputWialon"
for placa in placasWialon:
    for archivo in os.listdir(lugarDescargasWialon):
        archivo = lugarDescargasWialon + "\\" + archivo
        if placa in archivo:

            ###### Cronología ######
            try:
                wialonCronología = pd.read_excel(archivo, sheet_name="Cronología")
                wialonCronología = wialonCronología[wialonCronología['Tipo'] == 'Engine hours']
                wialonCronología['Placa'] = placa
                wialonCronología = wialonCronología['Placa'].value_counts().reset_index().rename(columns={"count":"Desplazamientos"})
            except:
                wialonCronología = pd.DataFrame({"Desplazamientos": "0", "Placa": placa}, index=[0])

            ###### Excesos de velocidad ######
            try:
                wialonExcesos = pd.read_excel(archivo, sheet_name="Excesos de velocidad")
                wialonExcesos['Placa'] = placa
                wialonExcesos = wialonExcesos[['Placa']].value_counts().reset_index().rename(columns={"count":"Excesos de velocidad"})
            except:
                wialonExcesos = pd.DataFrame({"Excesos de velocidad": "0", "Placa": placa}, index=[0])

            ###### Kilometraje ######
            try:
                wialonKilometraje = pd.read_excel(archivo, sheet_name="Statistics")
                wialonKilometraje = wialonKilometraje.T.drop('Unidad', axis=0)
                wialonKilometraje = wialonKilometraje[[6]].reset_index().rename(columns={"index":"Placa",6:"Kilometraje"})
            except:
                wialonKilometraje = pd.DataFrame({"Kilometraje": "0", "Placa": placa}, index=[0])

            ###### Merge por placa ######
            if placasWialon[0] == placa:
                wialonCompleto = wialonCronología.merge(wialonExcesos, on='Placa', how='outer')
                wialonCompleto = wialonCompleto.merge(wialonKilometraje, on='Placa', how="outer")
            else:
                wialonParcial = pd.merge(wialonCronología, wialonExcesos, on='Placa', how='outer').merge(wialonKilometraje, on='Placa', how="outer")
                wialonCompleto = pd.concat([wialonCompleto,wialonParcial])
        else:
            pass


####################################
###### Datos de placas Ubicom ######
####################################


# Desplazamientos
ubicomDesplazamientos = pd.read_excel(r"outputUbicom\Estacionados.xls", skiprows=16,usecols='F')
ubicomDesplazamientos = ubicomDesplazamientos[['Fecha']].value_counts().reset_index().rename(columns={"count":"Desplazamientos"})

# Excesos de velocidad y kilometraje
ubicomKilometrajeExcesos = pd.read_excel(r"outputUbicom\ReporteDiario.xls", skiprows=18, usecols='H,M,V')
ubicomKilometrajeExcesos = ubicomKilometrajeExcesos.drop(ubicomKilometrajeExcesos.tail(1).index,inplace=False).drop('Fecha',axis=1).rename(columns={"Distancia recorrida": "Kilometraje", "Número de excesos de velocidad": "Excesos de velocidad"})
ubicomKilometrajeExcesos['Fecha'] = ubicomDesplazamientos['Fecha']

ubicomCompleto = pd.merge(ubicomDesplazamientos,ubicomKilometrajeExcesos,on='Fecha',how="outer")
ubicomCompleto = ubicomCompleto.set_index('Fecha')
ubicomCompleto['Placa'] = "FNM236"
ubicomCompleto = ubicomCompleto.reset_index().drop('Fecha', axis=1).fillna(0)


####################################
###### Datos de placas Ubicar ######
####################################

lugarDescargasUbicar = r"C:\Users\pablo\Desktop\SGI - Práctica\RPA Vehículos\informesVehiculares\outputUbicar"
for archivo in os.listdir(lugarDescargasUbicar):
    archivo = lugarDescargasUbicar + "\\" + archivo
    if "stops" in archivo:
        # Desplazamientos
        ubicarDesplazamientos = pd.read_excel(archivo, usecols='A',skiprows=3)
        ubicarDesplazamientos = ubicarDesplazamientos.value_counts().loc[['Detenido']].reset_index()
        ubicarDesplazamientos = ubicarDesplazamientos.drop(columns="Unnamed: 0").rename(columns={"count":"Desplazamientos"})
        ubicarDesplazamientos['Placa'] = 'JYT620'
    else:
        # Kilometraje y excesos de velocidad
        ubicarKilometrajeExcesos = pd.read_excel(archivo, usecols='A,B',skiprows=2)
        ubicarKilometrajeExcesos = ubicarKilometrajeExcesos.loc[[0,5]].T.reset_index().drop(columns='index')
        ubicarKilometrajeExcesos = ubicarKilometrajeExcesos.drop(0,axis=0).rename(columns={0:"Kilometraje", 5:"Excesos de velocidad"})
        ubicarKilometrajeExcesos['Kilometraje'] = ubicarKilometrajeExcesos['Kilometraje'].str.extract(r'(\d+.\d+)').astype(float)
        ubicarKilometrajeExcesos['Placa'] = 'JYT620'

    # Merge final
ubicarCompleto = pd.merge(ubicarKilometrajeExcesos,ubicarDesplazamientos,on='Placa').fillna(0)


####################################
####### Datos de placas MDVR #######
####################################

lugarDescargasMDVR = os.getcwd() + "\\outputMDVR"
listaPlacasMDVR = ['KSZ298']
for placa in listaPlacasMDVR:
    for archivo in os.listdir(lugarDescargasMDVR):
        archivo = lugarDescargasMDVR + "\\" + archivo
        if "stops" in archivo:
            # Desplazamientos
            try:
                MDVRDesplazamientos = pd.read_excel(archivo, usecols='B', skiprows=3)
                MDVRDesplazamientos = MDVRDesplazamientos.value_counts().loc[['Movimiento']].reset_index()
                MDVRDesplazamientos = MDVRDesplazamientos.drop(columns="Unnamed: 1").rename(columns={"count":"Desplazamientos"})
                MDVRDesplazamientos['Placa'] = placa
            except:
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                wb = excel.Workbooks.Open(archivo)
                wb.SaveAs(archivo+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
                wb.Close()                               #FileFormat = 56 is for .xls extension
                excel.Application.Quit()
                if "xlsx" in archivo:
                    MDVRDesplazamientos = pd.read_excel(archivo, usecols='B', skiprows=3)
                    MDVRDesplazamientos = MDVRDesplazamientos.value_counts().loc[['Movimiento']].reset_index()
                    MDVRDesplazamientos = MDVRDesplazamientos.drop(columns="Unnamed: 1").rename(columns={"count":"Desplazamientos"})
                    MDVRDesplazamientos['Placa'] = placa
                else:
                    pass 
        else:
            # Kilometraje y excesos de velocidad
            MDVRKilometrajeExcesos = pd.read_excel(archivo, usecols='A,B',skiprows=2)
            MDVRKilometrajeExcesos = MDVRKilometrajeExcesos.loc[[0,5]].T.reset_index().drop(columns='index')
            MDVRKilometrajeExcesos = MDVRKilometrajeExcesos.drop(0,axis=0).rename(columns={0:"Kilometraje", 5:"Excesos de velocidad"})
            MDVRKilometrajeExcesos['Kilometraje'] = MDVRKilometrajeExcesos['Kilometraje'].str.extract(r'(\d+.\d+)').astype(float)
            MDVRKilometrajeExcesos['Placa'] = placa

# Merge final
MDVRCompleto = pd.merge(MDVRKilometrajeExcesos,MDVRDesplazamientos,on='Placa').fillna(0)


####################################
###### Datos de placas Ituran ######
####################################

archivoIturan = os.getcwd() + "\\" + r"outputIturan\Over speed by vehicle (summary).xls"
ituranCompleto = pd.read_excel(archivoIturan, usecols='C, D, J, K', skiprows=5, skipfooter=2)
ituranCompleto = ituranCompleto.rename(columns={"Ocurrencias":"Excesos de velocidad", " (Km)":"Kilometraje", "Viajes":"Desplazamientos"})


####################################
####### Merge de Plataformas #######
####################################


plataformasFinal = pd.concat([ubicomCompleto, ubicarCompleto, MDVRCompleto, securitracCompleto, wialonCompleto, ituranCompleto])
plataformasFinal = plataformasFinal.reset_index().drop('index',axis=1)
plataformasFinal = plataformasFinal.astype({'Kilometraje': float}).round({'Kilometraje': 2})
plataformasFinal['Desplazamientos'] = plataformasFinal['Desplazamientos'].apply(lambda x: "{:.2f}".format(float(x)))
plataformasFinal['Excesos de velocidad'] = plataformasFinal['Excesos de velocidad'].apply(lambda x: "{:.2f}".format(float(x)))
plataformasFinal = plataformasFinal.astype({'Desplazamientos': float, "Excesos de velocidad": float})
plataformasFinal.to_excel("plataformasFinal.xlsx", index=False)