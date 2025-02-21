import sys
import logging
from cli import CLI
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler

# Configurazione del logger
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

class Mega:
    def __init__(self):
        self.excel_handler = ExcelHandler()
        self.payload_handler = PayloadHandler(password="")

    def run(self):
        cli = CLI(self.excel_handler, self.payload_handler)
        cli.start()

if __name__ == '__main__':
    try:
        program = Mega()
        program.run()
    except Exception as e:
        logging.error(f"Errore nel programma: {e}")
        sys.exit(1)
