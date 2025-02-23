import time
from rich.console import Console
from rich.live import Live

console = Console()


class AsciiAnimation:
    def __init__(self, frames, delay=0.5, loop=False):
        """
        :param frames: lista di stringhe, ognuna rappresenta un frame dell'animazione
        :param delay: tempo di attesa tra un frame e l'altro (in secondi)
        :param loop: se True, l'animazione si ripete all'infinito
        """
        self.frames = frames
        self.delay = delay
        self.loop = loop
        self.colors = ["blue", "white", "red"]

    def play(self):
        color_index = 0
        with Live("", console=console, refresh_per_second=10) as live:
            while True:
                for frame in self.frames:
                    style = self.colors[color_index % len(self.colors)]
                    # Aggiorna il contenuto con il frame corrente formattato
                    live.update(f"[{style}]{frame}[/{style}]")
                    time.sleep(self.delay)
                    color_index += 1
                if not self.loop:
                    break