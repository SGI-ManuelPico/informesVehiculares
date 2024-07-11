import pandas as pd

tablaExcesos2 = pd.read_excel("pruebaCorreos.xlsx")
tablaExcesos2 = tablaExcesos2[tablaExcesos2['Conductor'] == "Pablo Ojeda"]

tablaExcesos2 = tablaExcesos2.set_index('Conductor')
print(tablaExcesos2.reset_index().iloc[0]['Conductor'])
    