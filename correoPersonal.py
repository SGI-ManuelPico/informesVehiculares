from util.correosVehiculares import CorreosVehiculares
import logging

def main():
    # Enviar correo al personal de SGI.
    try:
        CorreosVehiculares().enviarCorreoPersonal()
    except Exception as e:

        logging.error("Ocurrió un error", exc_info=True)


if __name__=='__main__':
    logging.basicConfig(level=logging.ERROR, format='%(asctime)s %(levelname)s %(message)s', filename='error.log')
    main()