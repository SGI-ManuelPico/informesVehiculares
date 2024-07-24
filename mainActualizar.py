import os
from db.consultasImportantes import ConsultaImportante
from persistence.actualizarIndividuales import ActualizarIndividuales
from itertools import permutations

def mainActualizarFaltantes():
    estado_plataforma = ConsultaImportante()
    campos_error = estado_plataforma.checkCamposError()
    
    for campo in campos_error:
        error_id = campo['id']
        plataforma = campo['plataforma']
        fecha = campo['fecha']

        current_dir = os.getcwd()
        ruta_error = os.path.join(current_dir, "SGI", plataforma, fecha)
        
        if os.path.exists(ruta_error) and os.listdir(ruta_error):
            # Tomar los archivos en la ruta especificada
            files = [os.path.join(ruta_error, f) for f in os.listdir(ruta_error)]
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


            def test_permutaciones(files, extraction_function):
                for perm in permutations(files):
                    result = extraction_function(*perm)
                    if result != []:
                        return perm  # Este es el orden correcto.
                return files  # Si ambas dan [] entonces una de dos: no hay nada para extraer (el carro no se movió) o la persona que bajo los archivos lo hizo mal (Una buena te pido Gina/Yuliana)

            
            
            actualizar = ActualizarIndividuales()
            if plataforma == 'Ituran':
                if len(plataforma_files['Ituran']) >= 2:
                    orden_correcto = test_permutaciones(plataforma_files['Ituran'], actualizar.extraerIturan)
                    actualizar.llenarIturan(orden_correcto[0], orden_correcto[1])
                    actualizar.llenarInfracIturan(orden_correcto[0], orden_correcto[1])
            elif plataforma == 'MDVR':
                if len(plataforma_files['MDVR']) >= 2: # Acá pueden pasar cosas feas porque el código de MDVR crea un nuevo archivo cuando se ejecuta. Pero esto se puede manejar.
                    orden_correcto = test_permutaciones(plataforma_files['MDVR'], actualizar.extraerMDVR)
                    actualizar.llenarMDVR(orden_correcto[0], orden_correcto[1])
                    actualizar.llenarInfracMDVR(orden_correcto[0], orden_correcto[1])
            elif plataforma == 'Securitrac': # Solo es un archivo 
                if len(plataforma_files['Securitrac']) >= 1:
                    actualizar.llenarSecuritrac(plataforma_files['Securitrac'][0])
                    actualizar.llenarInfracSecuritrac(plataforma_files['Securitrac'][0])
            elif plataforma == 'Ubicar':
                if len(plataforma_files['Ubicar']) >= 2:
                    orden_correcto = test_permutaciones(plataforma_files['Ubicar'], actualizar.extraerUbicar)
                    actualizar.llenarUbicar(orden_correcto[0], orden_correcto[1])
                    actualizar.llenarInfracUbicar(orden_correcto[0], orden_correcto[1])
            elif plataforma == 'Ubicom':
                if len(plataforma_files['Ubicom']) >= 2:
                    orden_correcto = test_permutaciones(plataforma_files['Ubicom'], actualizar.extraerUbicom)
                    actualizar.llenarUbicom(orden_correcto[0], orden_correcto[1])
            elif plataforma == 'Wialon':
                if len(plataforma_files['Wialon']) >= 3:
                    # No importa el orden
                    actualizar.llenarWialon(plataforma_files['Wialon'][0], plataforma_files['Wialon'][1], plataforma_files['Wialon'][2])
                    actualizar.llenarInfracWialon(plataforma_files['Wialon'][0], plataforma_files['Wialon'][1], plataforma_files['Wialon'][2])


            # Cambiar el estado a 'Gestionado'
            estado_plataforma.update_error_status(error_id, 'Gestionado')

if __name__ == "__main__":
    
    mainActualizarFaltantes()