import os
import sys

class AppSettings:
    """Configurazioni globali dell'applicazione"""

    # Ottieni path del progetto
    if getattr(sys, 'frozen', False):
        PROJECT_ROOT = os.path.dirname(sys.executable)
    else:
        PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
   
    
    # SAP Settings
    MAX_SAP_SESSIONS = 4
    SAP_CONNECTION_INDEX = 0
    SAP_TIMEOUT = 30
    
    # # Directory Settings
    # DATA_DIRECTORY = os.path.join(PROJECT_ROOT, "data")
    # DATA_DIRECTORY_SAP = os.path.join(DATA_DIRECTORY, "SAP")
    # DATA_DIRECTORY_EXCEL = os.path.join(DATA_DIRECTORY, "excel")

    # # Excel file Settings
    # MASK_TECHNOLOGY_EXCEL = "Mask_Tecnologia.xlsx"
    # ELENCO_IMPIANTI_EXCEL = "Elenco_Impianti.xlsx"
    # ELENCO_FL_EXCEL = "Elenco_FL.xlsx"
    # RESULT_EXCEL = "Result.xlsx"

    # GUI Settings
    WINDOW_WIDTH = 900
    WINDOW_HEIGHT = 600
    LOG_FONT = "Consolas"
    LOG_FONT_SIZE = 9
    
    # Threading Settings
    DEFAULT_WORKERS    = 4       # Numero di sessioni SAP parallele
    FL_TIMEOUT         = 60      # Secondi di attesa per singola FL (timeout)

    # Pausa periodica anti-saturazione SAP
    PAUSE_ENABLED      = False   # Attiva/disattiva la pausa
    PAUSE_BATCH_SIZE   = 1000    # FL per batch prima della pausa
    PAUSE_BASE_DELAY   = 60      # Durata base pausa in secondi
    PAUSE_PROGRESSIVE  = True    # True = pausa cresce ad ogni batch
    PAUSE_INCREMENT    = 30      # Secondi aggiuntivi per ogni batch successivo
    
    # # Tecnologie supportate
    # TECNOLOGIE = {
    #     'Wind': 'W',
    #     'Solar': 'S',
    #     'BESS': 'B',
    #     'Hydro': 'H'
    # }

    # # Maschere specifiche x tecnologia
    # TECH_MASK = {
    # 'H': "^([A-Z0-9]{2}H)(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{3})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2}))?)?)?)?)?$",
    # 'S': "^([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2}))?)?)?)?)?$",
    # 'W': "^([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2}))?)?)?)?)?$",
    # 'E': "^([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2}))?)?)?)?)?$"
    # }
