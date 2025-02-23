import pyfiglet
from rich.console import Console
from  rich.text import Text

from cli import Cli
from excel_handler import ExcelHandler
from payload_handler import PayloadHandler
from utils import CustomLogger
from animation import AsciiAnimation

logger = CustomLogger(__name__)
console = Console()

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
        # Definisci i frame della tua animazione ASCII

        frames = [
            r"""
ooo        ooooo 
`88.       .888'       
 888b     d'888   
 8 Y88. .P  888   
 8  `888'   888   
 8    Y     888   
o8o        o888o 
            """,
            r"""
ooo        ooooo oooooooooooo  
`88.       .888' `888'     `8  
 888b     d'888   888         
 8 Y88. .P  888   888oooo8    
 8  `888'   888   888    "    
 8    Y     888   888       o 
o8o        o888o o888ooooood8  
            """,
            r"""
ooo        ooooo oooooooooooo   .oooooo.      
`88.       .888' `888'     `8  d8P'  `Y8b        
 888b     d'888   888         888               
 8 Y88. .P  888   888oooo8    888              
 8  `888'   888   888    "    888     ooooo   
 8    Y     888   888       o `88.    .88'   
o8o        o888o o888ooooood8  `Y8bood8P'   
            """,
            r'''
ooo        ooooo oooooooooooo   .oooooo.          .o.       
`88.       .888' `888'     `8  d8P'  `Y8b        .888.      
 888b     d'888   888         888               .8"888.     
 8 Y88. .P  888   888oooo8    888              .8' `888.    
 8  `888'   888   888    "    888     ooooo   .88ooo8888.   
 8    Y     888   888       o `88.    .88'   .8'     `888.  
o8o        o888o o888ooooood8  `Y8bood8P'   o88o     o8888o  
            '''
        ]
        excel_icon = r'''
                 ++++++++++++++++===============
        ********************%%***+++++++++++++++
        ********************%%***+++++++++++++++
        ******  .***. .*****%%***+++++++++++++++
        *******. .*  -******%%***+++++++++++++++
        ********:   +*******%%***+++++++++++++++
        ********:   +******#%%###***************
        *******. =*  -**####%%###***************
        *****+  -***. .#####%%###***************
        *********###########%%###***************
              #**#####%%%%%%%%%%@@%#############
         '''
        animation = AsciiAnimation(frames, delay=0.8, loop=False)
        animation.play()
        banner_str = pyfiglet.figlet_format("Make Excel Great Again", font="small", width=150)
        banner_text = Text(banner_str, style="bold green")
        excel_icon_text = Text(excel_icon, style="bold green")
        console.print(banner_text, highlight=False, markup=False)
        console.print(excel_icon_text, justify="center", width=75, highlight=False, markup=False)
        # Avvia il programma principale
        program = Mega()
        program.run()

    except Exception as e:
        logger.error(f"Errore nel programma: {e}")
        exit(1)
