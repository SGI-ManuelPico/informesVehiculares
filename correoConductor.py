from util.correosVehiculares import CorreosVehiculares
import logging

def main():
    # Enviar correo específico a los conductores con excesos de velocidad.
    try:
        CorreosVehiculares().enviarCorreoConductor()
    except Exception as e:

        logging.error("Ocurrió un error", exc_info=True)



if __name__=='__main__':
    logging.basicConfig(level=logging.ERROR, format='%(asctime)s %(levelname)s %(message)s', filename='error.log')
    main()