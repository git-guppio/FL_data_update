# File: sap/operations.py

import time
import pandas as pd
import pyperclip
import win32clipboard
import re

from typing import List, Dict, Optional
from typing import Dict, Any, Optional, Tuple
from collections import Counter

import threading
from typing import Dict, Optional, List, Tuple
from .utils import SAPUtils
import os
from pathlib import Path
#from config.settings import AppSettings
from core.base_component import BaseComponent

class SAPOperations(BaseComponent):
    """Operazioni SAP specifiche per l'applicazione"""
    
    def __init__(self, logger=None):
        super().__init__(logger)
        self.sap_utils = SAPUtils(logger)
        
    def ricava_lista_impianti(self, session, tecnologia: str) -> Optional[str]:
        """
        Ricava la lista FL per una specifica tecnologia
        
        Args:
            session: Sessione SAP
            tecnologia: Codice tecnologia (W, S, B, H)
            
        Returns:
            Dati della clipboard o None
        """
        try:
            thread_id = threading.current_thread().ident

            self.log(f"[Thread {thread_id}] Ricavo lista FL: {tecnologia}", "info")
            
            # Vai alla transazione IH06
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nIH06"
            session.findById("wnd[0]").sendVKey(0)
            
            # Inserisci parametri
            session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = f"++{tecnologia}-++++"
            session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHECK_FL_S"

            # Inserisci filtro escludi
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").setFocus()
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey(2)
            time.sleep(0.25)
            # Selezione opzioni
            session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 5
            session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
            session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
            time.sleep(0.25)
            # Inserisci stato
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").text = "CRT"
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").caretPosition = 3
            session.findById("wnd[0]").sendVKey(0)

            # Avvia transazione
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.25)

            # Verifica status bar 
            msg_type, msg_text, _ = self.check_status_bar(session)
            
            if msg_type in ['E', 'ERROR']:
                self.log(f"❌ {tecnologia}: {msg_text}", "error")
                return {
                    'impianto': tecnologia,
                    'status': 'error',
                    'data': None,
                    'message': msg_text
                }
            
            # Esporta in clipboard
            session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select()
            time.sleep(0.5)
            
            # Seleziona formato clipboard
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            # Attendi completamento
            if not self.sap_utils.wait_for_sap(session, 30):
                self.log("Timeout durante esecuzione transazione", "error")
                return None
            
            time.sleep(0.5)
            
            # Attendi dati in clipboard
            if not self.sap_utils.wait_for_clipboard_data(30):
                self.log("Nessun dato trovato nella clipboard", "warning")
                return None
            
            # Leggi clipboard
            data = self.sap_utils.get_clipboard_data()
            
            # Torna al menu
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            
            self.log(f"[Thread {thread_id}] Lista FL ricavata con successo", "success")
            return data
            
        except Exception as e:
            thread_id = threading.current_thread().ident

            self.log(f"[Thread {thread_id}] Errore ricavo lista FL: {str(e)}", "error")
            return None
        finally:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]").sendVKey(0)
            except:
                pass

    def ricava_fl_impianto(self, session, impianto: str, data_directory: str) -> Optional[Dict]:
        """
        Ricava la lista FL per uno specifico impianto
        
        Args:
            session: Sessione SAP
            impianto: Codice impianto
            data_directory: Directory dove salvare i file
            
        Returns:
            Dizionario con risultato dell'operazione
        """                     

        try:
            thread_id = threading.current_thread().ident

            self.log(f"[Thread {thread_id}] Ricavo FL impianto: {impianto}", "info")
            
            # Vai alla transazione IH06
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nIH06"
            session.findById("wnd[0]").sendVKey(0)
            
            # Inserisci parametri
            session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = f"{impianto}" + "*"
            session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHECK_FL_PC"
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").setFocus()
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey(2)
            time.sleep(0.5)
            
            # Selezione opzioni
            session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellRow = 5
            session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
            session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
            time.sleep(0.5)
            
            # Inserisci stato
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").text = "CRT"
            session.findById("wnd[0]/usr/ctxtSTAE1-LOW").caretPosition = 3
            session.findById("wnd[0]").sendVKey(0)
            
            # Avvia transazione
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.5)
            # Verifica status bar
            msg_type, msg_text, _ = self.check_status_bar(session)
            
            if msg_type in ['E', 'ERROR']:
                self.log(f"❌ {impianto}: {msg_text}", "error")
                return {
                    'impianto': impianto,
                    'status': 'error',
                    'data': None,
                    'message': msg_text
                }
            # Caso in cui l'impianto esiste ma non ci sono sedi tecniche 
            elif msg_type in ['S', 'SUCCESS'] and msg_text == "Non sono stati selezionati oggetti":
                self.log(f"❌ {impianto}: {msg_text}", "error")
                return {
                    'impianto': impianto,
                    'status': 'error',
                    'data': None,
                    'message': msg_text
                }            
            
            # Esporta su file .csv
            session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.5)
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = data_directory + "\\"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{impianto}" + ".csv"
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            
            # Attendi completamento
            if not self.sap_utils.wait_for_sap(session, 30):
                self.log("Timeout durante esecuzione transazione", "error")
                return {
                    'impianto': impianto,
                    'data': None,
                    'status': 'error',
                    'errore': "Timeout durante esecuzione transazione"
                }
            
            time.sleep(0.5)

            # Verifica che l'impianto esista
            msg_type, msg_text, _ = self.check_status_bar(session)
            
            if msg_type in ['E', 'ERROR']:
                self.log(f"❌ {impianto}: {msg_text}", "error")
                return {
                    'impianto': impianto,
                    'status': 'error',
                    'data': None,
                    'message': msg_text
                }


            self.log(f"[Thread {thread_id}] Lista FL ricavata con successo", "success")
            return {
                'impianto': impianto,
                'data': impianto + ".csv",
                'status': 'success',
                'errore': None
            }
            
        except Exception as e:
            thread_id = threading.current_thread().ident

            self.log(f"[Thread {thread_id}] Errore ricavo lista FL: {str(e)}", "error")
            return {
                'impianto': impianto,
                'data': None,
                'status': 'error',
                'errore': str(e)
            }
        finally:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                session.findById("wnd[0]").sendVKey(0)
            except:
                pass

    def check_status_bar(self, session) -> Tuple[str, str, str]:
        """
        Verifica lo stato della Status Bar di SAP
        
        Args:
            session: Sessione SAP attiva
        
        Returns:
            Tuple (message_type, message_text, message_number)
        """
        try:
            status_bar = session.findById("wnd[0]/sbar")
            
            message_type = status_bar.MessageType.upper()
            message_text = status_bar.Text
            
            try:
                message_number = status_bar.MessageNumber
            except:
                message_number = ""
            
            icon = self.get_status_icon(message_type)
            self.log(f"Status Bar SAP: {icon} {message_type} - {message_text}", "debug")
            
            return message_type, message_text, message_number
            
        except Exception as e:
            self.log(f"Errore lettura Status Bar: {str(e)}", "error")
            return "E", f"Errore lettura status: {str(e)}", ""
    
    def verify_save_operation(self, session) -> bool:
        """
        Verifica che un'operazione di salvataggio sia andata a buon fine
        
        Args:
            session: Sessione SAP attiva
        
        Returns:
            bool: True se operazione riuscita
        """
        try:
            message_type, message_text, message_number = self.check_status_bar(session)
            
            if message_type in ['S', 'SUCCESS']:
                self.log(f"✅ Salvataggio riuscito: {message_text}", "success")
                return True
            
            if message_type in ['E', 'ERROR']:
                self.log(f"❌ Errore salvataggio: {message_text}", "error")
                return False
            
            if message_type in ['W', 'WARNING']:
                self.log(f"⚠ Warning salvataggio: {message_text}", "warning")
                return False
            
            if message_type in ['I', 'INFO']:
                self.log(f"ℹ Info salvataggio: {message_text}", "info")
                return False
            
            self.log(f"⚠ Tipo messaggio sconosciuto: {message_type}", "warning")
            return False
            
        except Exception as e:
            self.log(f"❌ Errore verifica salvataggio: {str(e)}", "error")
            return False            
    
    def get_status_icon(self, message_type: str) -> str:
        """
        Restituisce l'icona corrispondente al tipo di messaggio
        
        Args:
            message_type: Tipo messaggio ('S', 'W', 'E', 'I')
        
        Returns:
            str: Icona corrispondente
        """
        icons = {
            'S': '✅',
            'SUCCESS': '✅',
            'W': '⚠',
            'WARNING': '⚠',
            'E': '❌',
            'ERROR': '❌',
            'I': 'ℹ',
            'INFO': 'ℹ'
        }
        return icons.get(message_type.upper(), '❓')
    
