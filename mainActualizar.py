import os
from persistence.estadoPlataforma import EstadoPlataforma
from persistence.actualizarIndividuales import ActualizarIndividuales

def mainActualizarFaltantes():
    estado_plataforma = EstadoPlataforma()
    error_entries = estado_plataforma.check_for_error_entries()
    
    for entry in error_entries:
        error_id = entry['id']
        plataforma = entry['plataforma']
        fecha = entry['fecha']

        ''' Esto está bien feo de generalizar. Pero la idea es que la persona que descargue los archivos que falten los guarde, al menos por ahora, en una ruta como la que pongo de ejemplo.
         Lo que podemos hacer, es que cuando se registra el error en la tabla 'error', también se cree el directorio vacío para que la persona pueda meter los archivos ahí. 
         Entonces, lo que tocaría sería verificar en primera instancia si el directorio está vacio, y después verificar (*) si lo llenó con los archivos que eran.
        '''
        ruta_dummy = f"/SGI/{plataforma}/{fecha}/"
        
        '''     (*) ¿Cómo hacemos esta verificación?. Supongo que se podría hacer una extracción parcial y si eso funciona que siga. Lo que dificulta es el orden de los archivos.
                Si podemos asegurar que la persona que los guarda (Sea quien sea), lo haga con unos nomnbres especificos, entonces esto es más fácil.
                Si la ruta que queremos existe, entonces toca coger las rutas de los archivos que están en esta ruta y guardarlos. 
                Acá se me ocurren dos ideas. Podemos guardar las rutas que hay ruta_dummy en un diccionario y accederlas después como variables locales como hacemos en el otro main.
                La otra sería intentar accederlas de manera más directa con os, pero una vez más, ambas formas están atadas a que los nombres de los archivos sean siempre consistentes.
                Alternamente, y esto es bien chambón, como se sabe que son 3 archivos en el peor de los casos, entonces podemos decirle que pruebe ejecutando con los archivos en ordenes
                distintos y en el peor de los casos serían 6 pruebas porque solo se pueden permutar de 6 maneras.
                
        '''
        if os.path.exists(ruta_dummy) and os.listdir(ruta_dummy):
            # Tomar los archivos en la ruta especificada
            files = [os.path.join(ruta_dummy, f) for f in os.listdir(ruta_dummy)]
            files.sort()  # acá se ordenan en orden alfabetico, y después númerico. e.g: [ituran1, ituran2]

            # Acá se guardan las rutas para cada plataforma
            plataforma_files = {
                'Ituran': [],
                'MDVR': [],
                'Securitrac': [],
                'Ubicar': [],
                'Ubicom': [],
                'Wialon': []
            }

            # Agregamos las rutas a la llave (plataforma) correspondiente
            if plataforma in plataforma_files:
                plataforma_files[plataforma] = files

            # SOLO FUNCIONA SI LOS ARCHIVOS SE GUARDAN CON EL NOMBRE QUE QUEREMOS. POR EJEMPLO, SI PARA ITURAN GUARDAN LOS ARCHIVOS QUE DEBEN SER COMO ITURAN1, ITURAN2. Entonces 
            
            actualizar = ActualizarIndividuales()
            if plataforma == 'Ituran':
                if len(plataforma_files['Ituran']) >= 2:
                    actualizar.llenarIturan(plataforma_files['Ituran'][0], plataforma_files['Ituran'][1])
                    actualizar.llenarInfracIturan(plataforma_files['Ituran'][1])
            elif plataforma == 'MDVR':
                if len(plataforma_files['MDVR']) >= 2:
                    actualizar.llenarMDVR(plataforma_files['MDVR'][0], plataforma_files['MDVR'][1])
                    actualizar.llenarInfracMDVR(plataforma_files['MDVR'][1])
            elif plataforma == 'Securitrac':
                if len(plataforma_files['Securitrac']) >= 1: # Acá solo es un archivo entonces no hay lío.
                    actualizar.llenarSecuritrac(plataforma_files['Securitrac'][0])
                    actualizar.llenarInfracSecuritrac(plataforma_files['Securitrac'][0])
            elif plataforma == 'Ubicar':
                if len(plataforma_files['Ubicar']) >= 2:
                    actualizar.llenarUbicar(plataforma_files['Ubicar'][0], plataforma_files['Ubicar'][1])
                    actualizar.llenarInfracUbicar(plataforma_files['Ubicar'][0], plataforma_files['Ubicar'][1])
            elif plataforma == 'Ubicom': 
                if len(plataforma_files['Ubicom']) >= 2:
                    actualizar.llenarUbicom(plataforma_files['Ubicom'][0], plataforma_files['Ubicom'][1])
            elif plataforma == 'Wialon': # Esto funciona sin importar el orden.
                if len(plataforma_files['Wialon']) >= 3:
                    actualizar.llenarWialon(plataforma_files['Wialon'][0], plataforma_files['Wialon'][1], plataforma_files['Wialon'][2])
                    actualizar.llenarInfracWialon(plataforma_files['Wialon'][0], plataforma_files['Wialon'][1], plataforma_files['Wialon'][2])

            # Cambiar el estado a 'Gestionado'
            estado_plataforma.update_error_status(error_id, 'Gestionado')

if __name__ == "__main__":
    
    mainActualizarFaltantes()