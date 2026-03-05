import time
import win32clipboard
import pandas as pd
from typing import Optional
from io import StringIO
from core.base_component import BaseComponent

class SAPUtils(BaseComponent):
    """Utility per operazioni SAP comuni"""
    
    def __init__(self, logger=None):
        super().__init__(logger)     

    @staticmethod
    def wait_for_sap(session, timeout: int = 30) -> bool:
        """
        Attende che SAP finisca le operazioni in corso
        
        Args:
            session: Sessione SAP
            timeout: Tempo massimo di attesa in secondi
        
        Returns:
            bool: True se SAP è disponibile, False se timeout
        """
        start_time = time.time()
        
        try:
            while session.Busy:
                if time.time() - start_time > timeout:
                    print(f"Timeout dopo {timeout} secondi di attesa")
                    return False
                time.sleep(0.5)
            return True
        except Exception as e:
            print(f"Errore durante l'attesa: {str(e)}")
            return False
    
    @staticmethod
    def wait_for_clipboard_data(timeout: int = 30) -> bool:
        """
        Attende che la clipboard contenga dati
        
        Args:
            timeout: Tempo massimo di attesa in secondi
            
        Returns:
            bool: True se trovati dati, False se timeout
        """
        start_time = time.time()
        last_print_time = 0
        print_interval = 2
        
        while True:
            current_time = time.time()
            
            if current_time - start_time > timeout:
                print(f"Timeout: nessun dato dopo {timeout} secondi")
                return False
            
            try:
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        if data and data.strip():
                            print("Dati trovati nella clipboard")
                            return True
                finally:
                    win32clipboard.CloseClipboard()
                
                if current_time - last_print_time >= print_interval:
                    print("In attesa dei dati nella clipboard...")
                    last_print_time = current_time
                
                time.sleep(0.1)
                
            except win32clipboard.error as we:
                print(f"Errore Windows Clipboard: {str(we)}")
                time.sleep(0.5)
                continue
            except Exception as e:
                print(f"Errore controllo clipboard: {str(e)}")
                return False
    
    @staticmethod
    def get_clipboard_data() -> Optional[str]:
        """
        Legge dati dalla clipboard
        
        Returns:
            Stringa con il contenuto o None
        """
        try:
            win32clipboard.OpenClipboard()
            try:
                data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                return data if data else None
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"Errore lettura clipboard: {str(e)}")
            return None
    
    @staticmethod
    def clipboard_to_dataframe() -> Optional[pd.DataFrame]:
        """
        Converte il contenuto della clipboard in DataFrame
        
        Returns:
            DataFrame o None
        """
        try:
            data = SAPUtils.get_clipboard_data()
            if not data:
                return None
            
            # Converte la stringa in DataFrame
            df = pd.read_csv(StringIO(data), sep='\t')
            return df
        except Exception as e:
            print(f"Errore conversione clipboard a DataFrame: {str(e)}")
            return None