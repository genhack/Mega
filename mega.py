import sys
import time
import pyfiglet
from colorama import init, Fore, Back, Style
from cli import Cli
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler
from utils import CustomLogger

init(autoreset=True)
logger = CustomLogger(__name__)


class Mega:
    def __init__(self):
        # Passa il logger alle classi che lo richiedono
        self.excel_handler = ExcelHandler(logger)
        self.payload_handler = PayloadHandler(logger, password="")

    def run(self):
        # Passa il logger anche al CLI, insieme agli handler
        cli = Cli(logger, self.excel_handler, self.payload_handler)
        cli.start()

if __name__ == '__main__':
    try:
        # Genera l'ASCII art per "MEGA" e il sottotitolo
        mega_banner = pyfiglet.figlet_format("MEGA", font="big", width=150)
        subtitle_banner = pyfiglet.figlet_format("Make Excel Great Again", font="small", width=150)

        # Stampa i banner
        print(mega_banner)
        print(subtitle_banner)

        # Avvia il programma
        program = Mega()
        program.run()

    except Exception as e:
        logger.error(f"Errore nel programma: {e}")
        sys.exit(1)
