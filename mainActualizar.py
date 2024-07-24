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
                
        '''
        files_exist = os.path.exists(ruta_dummy)  

        if files_exist:
            # Perform individual update using ActualizarIndividuales class
            actualizar = ActualizarIndividuales()
            if plataforma == 'Ituran':
                actualizar.llenarIturan(ruta_dummy)
                actualizar.llenarInfracIturan(ruta_dummy)
            elif plataforma == 'MDVR':
                actualizar.llenarMDVR(ruta_dummy)
                actualizar.llenarInfracMDVR(ruta_dummy)
            elif plataforma == 'Securitrac':
                actualizar.llenarSecuritrac(ruta_dummy)
                actualizar.llenarInfracSecuritrac
            elif plataforma == 'Ubicar':
                actualizar.llenarUbicar(ruta_dummy)
                actualizar.llenarInfracUbicar
            elif plataforma == 'Ubicom':
                actualizar.llenarUbicom(ruta_dummy)
            elif plataforma == 'Wialon':
                actualizar.llenarWialon(ruta_dummy)
                actualizar.llnear

            # Actualizar el estado de esa actualización particular a especifica a 'Gestionado'
            estado_plataforma.update_error_status(error_id, 'Gestionado')

if __name__ == "__main__":
    
    mainActualizarFaltantes()