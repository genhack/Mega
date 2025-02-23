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
        if not self.filename:
            raise ValueError("Nessun file sorgente specificato.")
        if not self.filename.lower().endswith(self.extension_xlsx):
            raise ValueError(f"Il file '{self.filename}' non ha l'estensione {self.extension_xlsx}.")
        if not os.path.exists(self.filename):
            raise FileNotFoundError(f"Il file sorgente '{self.filename}' non esiste.")

        try:
            self.wb_source = xw.Book(self.filename)
            self.logging.info(f"File sorgente {self.filename} caricato con successo.")
        except Exception as e:
            raise RuntimeError(f"Errore durante il caricamento del file sorgente: {e}")

    def load_template(self) -> None:
        """
        Carica il file template in formato .xlsm.
        """
        if not self.template:
            raise ValueError("Nessun file template specificato.")
        if not self.template.lower().endswith(self.extension_xlsm):
            raise ValueError(f"Il file '{self.template}' non ha l'estensione {self.extension_xlsm}.")
        if not os.path.exists(self.template):
            raise FileNotFoundError(f"Il file template '{self.template}' non esiste.")

        try:
            self.wb_template = xw.Book(self.template)
            self.logging.info(f"Template {self.template} caricato con successo.")
        except Exception as e:
            raise RuntimeError(f"Errore durante il caricamento del template: {e}")

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
            # Verifica che i fogli esistano (meglio di accedere direttamente a sheets[0])
            if len(self.wb_source.sheets) == 0 or len(self.wb_template.sheets) == 0:
                raise RuntimeError("Il workbook non contiene fogli sufficienti.")
            src_sheet = self.wb_source.sheets[0]
            tmpl_sheet = self.wb_template.sheets[0]
            used_range = src_sheet.used_range
            used_range.copy(destination=tmpl_sheet.range("A1"))
            self.logging.info("Contenuto copiato in maniera 1:1 dal file sorgente al template.")
        except Exception as e:
            self.logging.error(f"Errore durante la copia del contenuto: {e}")
        finally:
            if self.wb_source:
                self.wb_source.close()

    def write_payload_to_cells(self, payload_b64: str, start_row: int = 2234, finish_row: int = 2235,
                               start_column: int = 177) -> None:
        """
        Scrive il payload Base64 nelle celle specificate del template.
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
            self.logging.info(f"Payload inserito con successo nelle celle del template: FU2234-FU2235")
        except Exception as e:
            raise RuntimeError(f"Errore durante la copia del payload: {e}")

    def save_workbook(self) -> None:
        """
        Salva il workbook template modificato in formato .xlsm, preservando le macro.
        """
        try:
            base, _ = os.path.splitext(self.filename)
            output_filename = base + "_mod" + self.extension_xlsm
            # Se il file esiste già, aggiungo un suffisso numerico per creare un nome univoco
            counter = 1
            new_filename = output_filename
            while os.path.exists(new_filename):
                new_filename = f"{base}_mod_{counter}{self.extension_xlsm}"
                counter += 1
            output_filename = new_filename
            self.logging.info(f"La cartella contiene un file con lo stesso nome, viene salvato come {output_filename}")

            self.wb_template.save(output_filename)
            self.wb_template.close()
            self.logging.info(f"File salvato come {output_filename}.")
        except Exception as e:
            self.logging.error(f"Errore durante il salvataggio del file: {e}")
