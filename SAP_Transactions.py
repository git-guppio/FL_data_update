import time
import pandas as pd
import pyperclip
import win32clipboard
import re

from typing import List, Dict, Optional
from typing import Dict, Any, Optional, Tuple
from collections import Counter


class SAPDataUpLoader:
    """ 
    Classe: SAPDataUpLoader
    Descrizione: Classe contenente i metodi per l' aggiornamento delle tabelle globali in SAP 
    """
    def __init__(self, session):
        """
        Inizializza la classe con una sessione SAP attiva
        
        Args:
            session: Oggetto sessione SAP attiva
        """
        self.session = session

    def update_table(self, table_name: str, data: pd.DataFrame) -> bool:
        """
        Aggiorna una tabella SAP con i dati forniti
        
        Args:
            table_name (str): Nome della tabella da aggiornare
            data (pd.DataFrame): Dati da inserire nella tabella
            
        Returns:
            bool: True se l'aggiornamento è riuscito, False altrimenti
        """
        try:

            return True
        except Exception as e:
            print(f"Errore durante l'aggiornamento della tabella {table_name}: {str(e)}")
            return False    


class SAPDataExtractor:
    """
    Classe per eseguire estrazioni dati da SAP utilizzando una sessione esistente
    """

    def __init__(self, session, main_window=None):
        self.session = session
        self.main_window = main_window
        # Configurazione messaggi multilingua
        self.SAP_MESSAGES = {
            'B_IH06_no_data_result': {
                'IT': "Non sono stati selezionati oggetti",
                'EN': "",
                'PT': "",
                'ES': ""
            },
            'W_IH06_multiple_data_result': {
                'IT': "Visualizzare sede tecnica: lista sedi tecniche",
                'EN': "",
                'PT': "",
                'ES': ""
            },
            'W_IH06_single_data_result': {
                'IT': "Visualizzare sede tecnica: Dati anagrafici",
                'EN': "",
                'PT': "",
                'ES': ""
            },
            'W_IFLO_selection_view': {
                'IT': "Data Browser: tabella IFLO: videata di selezione",
                'EN': "",
                'PT': "",
                'ES': ""
            },
            'W_IFLO_data_result': {
                'IT': r"Data Browser: tabella IFLO\s+\d+\s+hit",
                'EN': "",
                'PT': "",
                'ES': ""
            }                  
            # Aggiungi altri messaggi SAP qui...
        }    

    def check_sap_bar(self, message_bar: str, use_regex: bool = False) -> bool:
        """
        Verifica la presenza di un messaggio SAP nella lingua specificata
        
        Args:
            message_bar (str): Chiave del messaggio (es: 'data_browser_selection')
            lang (str): Codice lingua (es: 'IT', 'EN', 'DE')
            
        Returns:
            bool: True se il messaggio è trovato, False altrimenti
        """
        lang = self.main_window.infoLanguage
        try:
            window_bar = self.session.findById("wnd[0]/sbar").text
            # Verifica che il message_bar esista
            if message_bar not in self.SAP_MESSAGES:
                self.log_message(f"Message key '{message_bar}' non trovato", 'error')
                return False
            
            # Verifica che la lingua esista per questo messaggio
            messages = self.SAP_MESSAGES[message_bar]
            if lang not in messages:
                self.log_message(f"Lingua '{lang}' non supportata per '{message_bar}'. Lingue disponibili: {list(messages.keys())}", 'error')
                return False
            
            # Cerca il pattern nella lingua specifica
            expected_pattern = messages[lang]            
            
            # Verifica con regex o stringa normale
            if use_regex:
                match = re.search(expected_pattern, window_bar, re.IGNORECASE)
                if match:
                    self.log_message(f"✅ Pattern regex trovato in {lang}: '{expected_pattern}' -> '{match.group()}'", 'success')
                    return True
                else:
                    self.log_message(f"❌ Pattern regex non trovato in {lang}: '{expected_pattern}'", 'error')
                    return False            
            
            if expected_pattern in window_bar:
                self.log_message(f"✅ Finestra trovata in {lang}: {expected_pattern}", 'success')
                return True
            else:
                self.log_message(f"❌ Pattern non trovato in {lang}. Atteso: '{expected_pattern}', Trovato: '{window_bar}'", 'error')
                return False
            
        except Exception as e:
            self.log_message(f"Errore verifica finestra: {e}", 'error')
            return False

    def check_sap_window(self, message_key: str, use_regex: bool = False) -> bool:
        """
        Verifica la presenza di un messaggio SAP nella lingua specificata
        
        Args:
            message_key (str): Chiave del messaggio (es: 'data_browser_selection')
            lang (str): Codice lingua (es: 'IT', 'EN', 'DE')
            
        Returns:
            bool: True se il messaggio è trovato, False altrimenti
        """
        lang = self.main_window.infoLanguage
        try:
            window_text = self.session.findById("wnd[0]").text
            
            # Verifica che il message_key esista
            if message_key not in self.SAP_MESSAGES:
                self.log_message(f"Message key '{message_key}' non trovato", 'error')
                return False
            
            # Verifica che la lingua esista per questo messaggio
            messages = self.SAP_MESSAGES[message_key]
            if lang not in messages:
                self.log_message(f"Lingua '{lang}' non supportata per '{message_key}'.\nLingue disponibili: {list(messages.keys())}", 'error')
                return False
            
            # Cerca il pattern nella lingua specifica
            expected_pattern = messages[lang]

            # Verifica con regex o stringa normale
            if use_regex:
                match = re.search(expected_pattern, window_text, re.IGNORECASE)
                if match:
                    self.log_message(f"✅ Pattern regex trovato in {lang}: '{expected_pattern}' -> '{match.group()}'", 'success')
                    return True
                else:
                    self.log_message(f"❌ Pattern regex non trovato in {lang}: '{expected_pattern}'", 'error')
                    return False  
            
            if expected_pattern in window_text:
                self.log_message(f"✅ Finestra trovata in {lang}: {expected_pattern}", 'success')
                return True
            else:
                self.log_message(f"❌ Pattern non trovato in {lang}. \nAtteso: '{expected_pattern}', \nTrovato: '{window_text}'", 'error')
                return False
            
        except Exception as e:
            self.log_message(f"Errore verifica finestra: {e}", 'error')
            return False

    def log_message(self, message, icon_type='info'):
        """Wrapper per il log_message della main window"""
        if self.main_window:
            self.main_window.log_message(message, icon_type)
        else:
            print(message)  # Fallback su print

    def extract_FL_list(self, fl: str) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Estrae la lista delle FL 
        
        Args:
            fl (str): Codice Functional Location
            
        Returns:
            Tuple[bool, Optional[Dict[str, Optional[str]]]]: 
                - bool: True se estrazione riuscita, False altrimenti
                - df: dataframe contenente le informazioni estratte
        """
        try:
            # Utilizza transazione IH06
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nIH06"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = fl
            self.session.findById("wnd[0]/usr/ctxtVARIANT").text = "CHECK_FL_S"
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            # attendo il caricamento dei dati
            time.sleep(0.5)
            ## Verifico se sono stati trovati dati
            # Nessun dato travato
            if self.check_sap_bar('B_IH06_no_data_result'):
                raise ValueError(f"Nessun dato per la FL: {fl}")
            #  Un solo valore trovato
            elif self.check_sap_window('W_IH06_single_data_result'):
                self.log_message(f"Numero di elementi per la FL {fl} = 1", "info")
                # Creo il df ed inserisco il valore della FL
                df_fl = pd.DataFrame({"Sede tecnica": [fl]})
                # Leggo il valore della definizione sede tecnica e lo inserisco nel df
                definizione = self.session.findById("wnd[0]/usr/txtIFLO-PLTXT").text
                df_fl["Definizione della sede tecnica"] = definizione
                return True, df_fl
            # Più di un valore trovato
            elif self.check_sap_window('W_IH06_multiple_data_result'):
                num_elementi = self.session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount
                self.log_message(f"Numero di elementi per la FL {fl} = {num_elementi}", "info")
                self.session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select()
                time.sleep(0.5)  
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                # Attendi che SAP sia pronto
                time.sleep(0.5)
                # Attendi che la clipboard sia riempita
                if not self.wait_for_clipboard_data(30):
                    # Gestisci il caso in cui non sono stati trovati dati
                    print("Nessun dato trovato nella clipboard")
                    # Eventuali azioni di fallback
                # Leggo il contenuto della clipboard
                fl_data = self.clipboard_data()
                if fl_data is None:
                    raise ValueError(f"Nessun dato presente nella clipboard")
                result, df_fl = self.clean_data(fl_data)
                if not result:
                    raise ValueError(f"Errore durante la pulizia dei dati della FL {fl}")
                else:
                    return True, df_fl
        except Exception as e:
            self.log_message(f"Errore durante l'estrazione delle informazioni da FL {fl}: \n{str(e)}")
            return False, None

    def extract_FL_IFLO(self, d_fl: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Estrae la lista delle FL 
        
        Args:
            d_fl: dataframe contenente le FL da estrarre
            
        Returns:
            Tuple[bool, pd.DataFrame]: 
                - bool: True se estrazione riuscita, False altrimenti
                - df: dataframe contenente le informazioni estratte
        """
           # copio i dati contenuti nel df nella clipboard
        if not self.copy_values_for_sap_selection(d_fl):
            return False, None
           # Se la copia dei dati è andata a buon fine, procedo con l'estrazione
        try:
            # Avvio transazione SE16
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16"
            self.session.findById("wnd[0]").sendVKey(0)
            # Richiedo tabella IFLO
            self.session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "IFLO"
            self.session.findById("wnd[0]").sendVKey(0)
            # Attendo il caricamento della tabella
            time.sleep(0.5)
            # verifico il titolo della finestra
            if not self.check_sap_window('W_IFLO_selection_view'):
                self.log_message("Errore: la tabella IFLO non è stata trovata", "error")
                raise ValueError("Tabella IFLO non trovata")
            # Apro finestra per inserimento valori FL
            self.session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
            # Copio valori da Clipboard
            self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            # attendo il caricamento dei dati
            time.sleep(0.25)
            # Verifico che i dati siano stati copiati (almeno un valore nella finestra di testo)
            if self.session.findById("wnd[0]/usr/ctxtI1-LOW").text == "":
                self.log_message("Nessun valore inserito per la FL", "error")
                raise ValueError("Nessun valore inserito per la FL")
            # Seleziono lingua principale
            self.session.findById("wnd[0]/usr/txtI4-LOW").text = "X"
            # Modifico n. massimo risultati
            self.session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
            # Avvio la transazione
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            # Attendo il caricamento dei dati
            time.sleep(0.5)
            # Verifico che siano stati trovati dati leggendo il nome della finestra
            if self.check_sap_window('W_IFLO_data_result', True):
                # Se non trova il pattern, allora verifico se è presente un icona di errore nella status bar
                try:
                    iconType = self.session.findById("wnd[0]/sbar").MessageType
                    if iconType == 'E': # dovrebbe essere indipendente dalla lingua
                        self.log_message("FL inesistenti", "error")
                        raise ValueError("FL selezionate inesistenti")
                except AttributeError:
                    # Se l'oggetto non ha l'attributo MessageType, gestisco l'errore
                    self.log_message("Errore: impossibile leggere il tipo di icona nella status bar", "error")
                    return False, None 
            ### La finestra aperta è corretta 
            # Apro il menu per la selezione del template
            self.session.findById("wnd[0]/tbar[1]/btn[33]").press()
            ### Ricerco il template nell'elenco
            # Riferimento alla griglia
            grid = self.session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
            # Parametri di ricerca
            target_value = "CHECK_FL_L"
            row_count = grid.RowCount
            layout_ok = False
            # Ricerca del valore nella prima colonna
            for i in range(row_count):
                try:
                    cell_value = grid.getCellValue(i, grid.ColumnOrder(0))
                    if cell_value == target_value:
                        print(f"Valore trovato alla riga: {i}")
                        grid.currentCellRow = i
                        grid.selectedRows = str(i)
                        grid.clickCurrentCell()
                        layout_ok = True
                        break
                    
                except Exception as e:
                    print(f"Errore nella selezione del layout {i}: {e}")
                    continue
            if not layout_ok:
                # Se il layout non è stato trovato, gestisco l'errore
                self.log_message(f"Layout '{target_value}' non trovato nella griglia", "error")
                return False, None
            else:
                # verifico l'icona che compare nella status bar
                # Il valore restituito dovrebbe indicare il tipo di icona mostrata:
                #     - 'S' o 'SUCCESS' per il simbolo di successo (✓)
                #     - 'W' o 'WARNING' per l'icona di avviso (⚠)
                #     - 'E' o 'ERROR' per l'icona di errore (❌)
                #     - 'I' o 'INFO' per l'icona informativa (ℹ)
                try:
                    iconType = self.session.findById("wnd[0]/sbar").MessageType
                    if iconType != 'S':
                        self.log_message("Errore nella selezione del Layout", "error")
                        return False, None
                except AttributeError:
                    # Se l'oggetto non ha l'attributo MessageType, gestisco l'errore
                    self.log_message("Errore: impossibile leggere il tipo di icona nella status bar", "error")
                    return False, None      
                except Exception as e:
                    self.log_message(f"Errore durante la lettura del tipo di icona nella status bar: {str(e)}", "error")
                    return False, None
            
            ### Se la selezione del layout è andata a buon fine, copio i dati nella clipboard
            self.session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()            
            # attendo il caricamento dei dati
            time.sleep(0.5)
            # Leggo il contenuto della clipboard
            fl_data = self.clipboard_data()
            if fl_data is None:
                raise ValueError(f"Nessun dato presente nella clipboard")
            result, df_fl = self.clean_data(fl_data)
            if not result:
                raise ValueError(f"Errore durante la pulizia dei dati della FL {fl}")
            else:
                return True, df_fl
        
        except Exception as e:
            self.log_message(f"Errore durante l'estrazione delle informazioni da FL:\n{str(e)}")
            return False, None

    def modify_FL(self, fl: str, descrizione: str) -> Tuple[bool, Optional[str]]:
        """
        Modifica le informazioni della Functional Location
        Args:
            fl (str): Codice Functional Location
            descrizione (str): Nuova descrizione della Functional Location
            
        Returns: 
                - bool: True se estrazione riuscita, False altrimenti
                - str: Messaggio di successo o errore
        """
        try:

            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nIL02"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            # Inserisco la FL da modificare
            self.session.findById("wnd[0]/usr/ctxtIFLO-TPLNR").text = fl
            # Avvio transazione
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            # inserisco descrizione
            self.session.findById("wnd[0]/usr/txtIFLO-PLTXT").text = descrizione
            self.session.findById("wnd[0]").sendVKey(0)
            # Salvo i dati
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(0.5)
            # Verifico icona della status bar
                # verifico l'icona che compare nella status bar
                # Il valore restituito dovrebbe indicare il tipo di icona mostrata:
                #     - 'S' o 'SUCCESS' per il simbolo di successo (✓)
                #     - 'W' o 'WARNING' per l'icona di avviso (⚠)
                #     - 'E' o 'ERROR' per l'icona di errore (❌)
                #     - 'I' o 'INFO' per l'icona informativa (ℹ)
            try:
                iconType = self.session.findById("wnd[0]/sbar").MessageType
                if iconType != 'S':
                    self.log_message("Errore nella modifica FL {fl}", "error")
                    return False, "Errore nella modifica della FL"
                else:
                    return True, "FL modificata"     
            except Exception as e:
                self.log_message(f"Errore durante la lettura del tipo di icona nella status bar: {str(e)}", "error")
                return False, e       
   
        except Exception as e:  
            self.log_message(f"Errore durante la modifica della FL {fl}: \n{str(e)}")
            return False, f"Errore: {str(e)}"
        
#-----------------------------------------------------------------------------
# Metodi per la gestione della clipboard
#-----------------------------------------------------------------------------

    def clean_data(self, data) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Legge i dati dalla clipboard, rimuove le righe di separazione e le colonne vuote,
        e gestisce le intestazioni duplicate.
        
        Returns:
            DataFrame Pandas pulito o False in caso di errore
        """
        try:
            # Controlla se i dati sono presenti
            if not data:
                raise ValueError(f"Nessun dato trovato")
            # Controlla se i dati sono sufficienti
            all_lines = data.strip().split('\n')
            if len(all_lines) <= 3:
                raise ValueError(f"Il file deve contenere almeno 4 righe, trovate solo {len(all_lines)}")      

            # Divide in righe
            lines = all_lines[2:]  # Salta le prime tre righe che sono intestazioni
            
            # Filtra le righe, escludendo quelle che contengono solo trattini
            clean_lines = []
            for line in lines:
                # Rimuove spazi bianchi iniziali e finali
                line = line.strip()
                # Verifica se la riga è composta solo da trattini
                if line and not all(c == '-' for c in line.replace(' ', '')):
                    clean_lines.append(line)

            if not clean_lines:
                print("Nessuna riga valida trovata dopo la pulizia")
                return False, None

            # Dividi le righe in colonne usando il tab come separatore
            data_rows = [line.split('|') for line in clean_lines]
            
            # Prendi la prima riga come header
            original_headers = [header.strip() for header in data_rows[0]]
            
            # Gestisci gli header duplicati
            unique_headers = self.handle_duplicate_headers(original_headers)
            
            # Se sono stati trovati duplicati, stampalo
            duplicates = [header for header, count in Counter(original_headers).items() if count > 1]
            if duplicates:
                print("\nTrovate colonne con nomi duplicati:")
                for dup in duplicates:
                    print(f"- '{dup}' (rinominate con postfissi numerici)")

            # Crea il DataFrame con i nuovi header
            df = pd.DataFrame(data_rows[1:], columns=unique_headers)

            # Rimuove le colonne completamente vuote o con tenenti valori nulli
            cols_to_keep = []
            for col in df.columns:
                col_clean = df[col].astype(str).str.strip()
                if not col_clean.isin(['', 'nan', 'None', 'NaN']).all():
                    cols_to_keep.append(col)
            if not cols_to_keep:
                raise ValueError("Nessuna colonna contiene dati validi")
                
            df = df[cols_to_keep]
            print(f"✅ DataFrame filtrato: {len(cols_to_keep)} colonne mantenute")
            # Verifico se il df contiene dei dati
            if df.empty:
                return False, None
            # Reset dell'indice
            df = df.reset_index(drop=True)
         
            return True, df
        
        except Exception as e:
            print(f"Errore durante la pulizia dei dati: {str(e)}")
            return False, None
           
    def handle_duplicate_headers(self, headers: List[str]) -> List[str]:
        """
        Gestisce le intestazioni duplicate aggiungendo un postfisso numerico
        
        Args:
            headers: Lista delle intestazioni originali
            
        Returns:
            Lista delle intestazioni con postfissi per i duplicati
        """
        # Conta le occorrenze di ogni header
        header_counts = Counter()
        unique_headers = []
        
        for header in headers:
            # Se l'header è già stato visto
            if header in header_counts:
                # Incrementa il contatore e aggiungi il postfisso
                header_counts[header] += 1
                unique_headers.append(f"{header}_{header_counts[header]}")
            else:
                # Prima occorrenza dell'header
                header_counts[header] = 0
                unique_headers.append(header)
        
        return unique_headers

    def wait_for_clipboard_data(self, timeout: int = 30) -> bool:
        """
        Attende che la clipboard contenga dei dati
        
        Args:
            timeout: Tempo massimo di attesa in secondi
            
        Returns:
            bool: True se sono stati trovati dati, False se è scaduto il timeout
        """
        start_time = time.time()
        last_print_time = 0  # Per limitare i messaggi di log
        print_interval = 2   # Intervallo in secondi tra i messaggi di log
        
        while True:
            current_time = time.time()
            
            # Verifica timeout
            if current_time - start_time > timeout:
                print(f"Timeout: nessun dato trovato nella clipboard dopo {timeout} secondi")
                return False
            
            try:
                # Controlla il contenuto della clipboard
                win32clipboard.OpenClipboard()
                try:
                    # Verifica se c'è del testo nella clipboard
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        if data and data.strip():
                            print("Dati trovati nella clipboard")
                            return True
                finally:
                    win32clipboard.CloseClipboard()
                
                # Stampa il messaggio di attesa solo ogni print_interval secondi
                if current_time - last_print_time >= print_interval:
                    print("In attesa dei dati nella clipboard...")
                    last_print_time = current_time
                
                # Aspetta prima del prossimo controllo
                time.sleep(0.1)  # Ridotto il tempo di attesa per una risposta più veloce
                
            except win32clipboard.error as we:
                print(f"Errore Windows Clipboard: {str(we)}")
                time.sleep(0.5)  # Attesa più lunga in caso di errore
                continue
            except Exception as e:
                print(f"Errore durante il controllo della clipboard: {str(e)}")
                return False  

    def clipboard_data(self) -> Optional[str]:
        """
        Legge i dati dalla clipboard.
        
        Returns:
            DataFrame Pandas pulito o None in caso di errore
        """
        try:
            # Legge il contenuto della clipboard
            win32clipboard.OpenClipboard()
            try:
                data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()

            if not data:
                print("Nessun dato trovato nella clipboard")
                return None
            else:
                 return data

        except Exception as e:
            print(f"Errore durante lettura dei dati dalla clipboard: {str(e)}")
            return None        
        
    def copy_values_for_sap_selection(self, values: pd.DataFrame) -> bool:
        """
        Copia valori formattati nella clipboard per utilizzarli in un campo di selezione multipla SAP.
        
        Args:
            values: DataFrame
        """
        try:
            # Gestione DataFrame pandas
            if isinstance(values, pd.DataFrame):
                if values.empty:
                    self.log_message("Nessun valore da copiare", "warning")
                    return False
                # Estrai valori dal DataFrame
                values_list = values.values.flatten().tolist()
            
            # Filtra valori validi
            valid_values = [str(value) for value in values_list if pd.notna(value) and str(value).strip()]
            
            # Converte la lista in una stringa per la clipboard
            text = '\r\n'.join(valid_values)
            
            # Conta le righe nella stringa
            num_righe = len(text.split('\r\n')) if text else 0
            
            # Copia nella clipboard
            pyperclip.copy(text)
            time.sleep(0.1)
            
            # Log con informazioni sui valori copiati
            self.log_message(f"Copiati {num_righe} valori nella clipboard per SAP", "success")
            return True
            
        except Exception as e:
            self.log_message(f"Errore durante la copia nella clipboard: {str(e)}", "error")
            return False