import sys
from cli import Cli
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler
from utils import CustomLogger

logger = CustomLogger(__name__)

class Mega:
    def __init__(self):
        # Passa il logger alle classi che lo richiedono
        self.excel_handler = ExcelHandler(logger)
        self.payload_handler = PayloadHandler(logger, password="")  # Assicurati di passare il logger come primo parametro

    def run(self):
        # Passa il logger anche al CLI, insieme agli handler
        cli = Cli(logger, self.excel_handler, self.payload_handler)
        cli.start()

if __name__ == '__main__':
    try:
        program = Mega()
        program.run()
    except Exception as e:
        logger.error(f"Errore nel programma: {e}")
        sys.exit(1)
