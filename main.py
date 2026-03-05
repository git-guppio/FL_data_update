import re
from pathlib import Path
import os
import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                           QHBoxLayout, QWidget, QTextEdit, QListWidget, QLabel, QMessageBox,
                           QDialog, QRadioButton, QButtonGroup, QDialogButtonBox, QMenu, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QTextCursor

from sap.session_manager import SAPSessionManager
from sap.operations import SAPOperations
from core.thread_manager import ThreadManager
from core.base_component import BaseComponent # Importa la classe per utilizzo omogeneo del logging
from core.logger import ThreadSafeLogger
#import SAP_Connection
from SAP_Transactions import SAPDataExtractor
from typing import Tuple, Optional, Dict

from config.settings import AppSettings

# import logging

# # Configurazione base del logging per tutta l'applicazione
# logging.basicConfig(
#     level=logging.INFO,
#     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
#     handlers=[
#         logging.FileHandler("app.log"),
#         logging.StreamHandler()
#     ]
# )

# # Logger specifico per questo modulo
# logger = logging.getLogger("main").setLevel(logging.DEBUG)

class MainWindow(QMainWindow, BaseComponent):
    def __init__(self):
        #super().__init__()
        QMainWindow.__init__(self)
        logger = ThreadSafeLogger() # Creo un logger thread-safe
        BaseComponent.__init__(self, logger) # Inizializzo senza logger, lo assegno dopo        

        
        # Ottiene il percorso della directory del file Python corrente
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        # Inizializza variabili per memorizzare informazioni sulla connessione SAP
        self.infoUser = ""
        self.infoSystemName = ""
        self.infoClient = ""
        self.infoLanguage = ""
        # Pattern per la verifica delle FL inserite
        self.patterns = {
            # 'MaskGenerica': r'^(?:([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2,3})(?:-([A-Z0-9]{2,3})(?:-([A-Z0-9]{2}))?)?)?)?)?)?$',
            'Mask_gen': r'^(?:([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2,3})(?:-([A-Z0-9]{2,3})(?:-([A-Z0-9]{2}))?)?)?)?)?)?$',
            'Mask_star': r'^(?:([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:[A-Z0-9*\-]{1,13}))?)?$'
            # aggiungere altre maschere se necessario
        }
        self.fl_dictionary = {} # Dizionario per memorizzare le FL dalla finestra di testo a sx
        self.fl_df_tot = pd.DataFrame()  # DataFrame per memorizzare tutti i dati estratti

        # Inizializza componenti per il MultiThreaD
        self.session_manager = None
        self.thread_manager = None # Inizializzato in initialize_sap_components()
        self.sap_operations = None      

        # Setup GUI
        self.init_ui()          

        # Setup timer per processare log dalla coda
        self.log_timer = QTimer()
        self.log_timer.timeout.connect(self.process_log_queue)
        self.log_timer.start(100)  # Controlla ogni 100ms


    def init_ui(self):
        # Inizializza l'interfaccia utente
        self.setWindowTitle("Aggiorna valori FL")
        self.setGeometry(100, 100, 1000, 600)
        # Widget centrale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principale
        main_layout = QVBoxLayout(central_widget)
        
        # Layout orizzontale per i due pannelli
        content_layout = QHBoxLayout()
        
        # Pannello sinistro (TextEdit per clipboard)
        left_panel = QVBoxLayout()
        left_label = QLabel("Dati da Clipboard:")
        left_panel.addWidget(left_label)
        
        self.clipboard_area = QTextEdit()
        self.clipboard_area.setPlaceholderText("Inserire Parent da cui iniziare ricorsivamente l'aggiornamento delle FL\n"
                                                "Esempio: \nESS-ESND\nESS-ESSW-52\n")

        left_panel.addWidget(self.clipboard_area)
        
        # Aggiungi il layout sinistro al layout orizzontale
        content_layout.addLayout(left_panel)
        
        # Pannello destro (ListView per log)
        right_panel = QVBoxLayout()
        right_label = QLabel("Log operazioni:")
        right_panel.addWidget(right_label)

        # Area di testo per il log
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont(AppSettings.LOG_FONT, AppSettings.LOG_FONT_SIZE))
        
        # Configura il formato del blocco per aumentare lo spacing
        cursor = self.log_text.textCursor()
        block_format = cursor.blockFormat()
        block_format.setLineHeight(110, 1)  # % dell'altezza normale (tipo 1 = percentuale)
        cursor.setBlockFormat(block_format)
        self.log_text.setTextCursor(cursor)

        right_panel.addWidget(self.log_text)

        # Attiva il menu contestuale per il widget dei log
        self.log_text.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log_text.customContextMenuRequested.connect(self.show_context_menu)        
        
        # Aggiungi il layout destro al layout orizzontale
        content_layout.addLayout(right_panel)
        
        # Aggiungi il layout dei contenuti al layout principale
        main_layout.addLayout(content_layout)
        
        # Layout per i bottoni
        button_layout = QHBoxLayout()
        
        # Bottone Pulisci
        self.clear_button = QPushButton('Pulisci Finestre')
        self.clear_button.clicked.connect(self.clear_windows)
        button_layout.addWidget(self.clear_button)
        
        # Bottone Estrai
        self.extract_button = QPushButton('Aggiorna Dati')
        self.extract_button.clicked.connect(self.update_data)
        button_layout.addWidget(self.extract_button)
        
        # Bottone Upload
        self.upload_button = QPushButton('Salva Dati')
        self.upload_button.clicked.connect(self.save_data)
        self.upload_button.setEnabled(False)  # Disabilitato finché non implementato
        button_layout.addWidget(self.upload_button)
        
        # Aggiungi il layout dei bottoni al layout principale
        main_layout.addLayout(button_layout)

    def process_log_queue(self):
        """
        Processa i messaggi dalla coda thread-safe
        
        Questo è l'UNICO metodo che scrive nel QTextEdit
        Viene chiamato dal timer ogni 100ms
        """
        messages = self.logger.process_queue()
        
        for timestamp, message, msg_type in messages:
            formatted_message = f"[{timestamp}] {message}"
            color = ThreadSafeLogger.get_color_for_type(msg_type)
            
            # Scrivi nel widget
            self.log_text.setTextColor(color)
            self.log_text.append(formatted_message)
        
        if messages:
            # Scroll automatico
            cursor = self.log_text.textCursor()
            cursor.movePosition(QTextCursor.End)
            self.log_text.setTextCursor(cursor)        
    
    # ----------------------------------------------------
    # Funzioni per mostrare un menu contestuale x copiare i dati
    # ----------------------------------------------------
    def show_context_menu(self, position):
        # Crea menu contestuale
        context_menu = QMenu()
        
        # Aggiungi l'azione "Copia selezione" (solo se c'è testo selezionato)
        if self.log_text.textCursor().hasSelection():
            copy_action = QAction("Copia selezione", self)
            copy_action.triggered.connect(self.copy_selected_items)
            context_menu.addAction(copy_action)
        
        # Aggiungi l'azione "Copia tutto"
        copy_all_action = QAction("Copia tutto", self)
        copy_all_action.triggered.connect(self.copy_all_items)
        context_menu.addAction(copy_all_action)
        
        # Mostra il menu contestuale alla posizione corrente del cursore
        context_menu.exec_(QCursor.pos())

    def copy_selected_items(self):
        # Usa il metodo copy() integrato di QTextEdit
        if self.log_text.textCursor().hasSelection():
            self.log_text.copy()  # Copia automaticamente nella clipboard
            print("Testo selezionato copiato negli appunti")
        else:
            print("Nessun testo selezionato")

    def copy_all_items(self):
        # Copia tutto il contenuto del QTextEdit
        all_text = self.log_text.toPlainText()
        
        # Controlla se c'è testo da copiare
        if all_text.strip():  # Verifica che non sia vuoto o contenga solo spazi
            QApplication.clipboard().setText(all_text)
            print("Tutto il testo copiato negli appunti")
        else:
            print("Nessun testo da copiare")    


    def log_message(self, message, icon_type='info'):
        """
        Wrapper per SAPDataExtractor: delega a self.log() che usa il ThreadSafeLogger.
        """
        self.log(message, icon_type)

    # def log_message(self, message, icon_type='info'):
    #     """
    #     Aggiunge un messaggio al log con un'icona Qt
    #     """
    #     item = QListWidgetItem(message)
        
    #     # Imposta l'icona in base al tipo
    #     if icon_type == 'info':
    #         item.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
    #     elif icon_type == 'error':
    #         item.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxCritical))
    #     elif icon_type == 'success':
    #         item.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
    #     elif icon_type == 'warning':
    #         item.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxWarning))
    #     elif icon_type == 'loading':
    #         item.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        
    #     self.log_text.addItem(item)
    #     self.log_text.scrollToBottom()


    """ 
        def log_message(self, message, icon_type='info'):
            
            #Aggiunge un messaggio al log con un'emoji come icona
            

            icons = {
                'info': '\U0001f604',
                'error': '❌',
                'success': '✅',
                'warning': '⚠️',
                'loading': '⏳'
            }  
            icon = icons.get(icon_type, '')
            self.log_text.addItem(f"{icon} {message}")
            self.log_text.scrollToBottom()
    """    

    def clear_windows(self):
        self.clipboard_area.clear()
        self.log_text.clear()
        self.extract_button.setEnabled(True)
        self.upload_button.setEnabled(False)
        self.log("Finestre pulite")
        # Elimino i dati memorizzati da estrazioni precedenti
        self.fl_dictionary = {}
        self.fl_df_tot = pd.DataFrame()

    def validate_clipboard_data(self) -> Tuple[bool, dict[str, pd.DataFrame] | None]:
        """Valida i dati nella finestra di testo sinistra (clipboard_area)"""
        data = self.clipboard_area.toPlainText().strip().split('\n')
        data = [line.strip() for line in data if line.strip()]  # Rimuove linee vuote
        
        # Verifica se ci sono dati
        if not data:
            QMessageBox.warning(self, "Attenzione", "Inserire i dati nella finestra di sinistra prima di procedere.")
            return False
        # Compila i pattern per la validazione
        patterns = self.patterns
        fl_dictionary = {}
        fl_errors = ""
        #lines = data.split('\n')
        for i, line in enumerate(data, 1):
            if not line.strip():
                continue
            # Data contiene le righe presenti nella clipboard_area (riquadro a sx)
            # Le righe possono contenere codici di sedi tecniche complete oppure dei codici contenenti il carattere '*'
            # Nel primo caso verifico che la riga rispetti la maschera 'Mask_gen' e inserisco le riga all'interno del df fl_dictionary['Mask_gen']
            # Nel secondo caso verifico che la riga rispetti la maschera 'Mask_star' e creo una nuova chiave nel dizionario che andrà a contenere le FL estratte con transazione H06
            try:
                if  ('*' not in line) and (re.match(patterns['Mask_gen'], line)):
                    # Verifica se la chiave esiste già nel dizionario
                    if 'Mask_gen' not in fl_dictionary:
                        fl_dictionary['Mask_gen'] = pd.DataFrame()
                    # Aggiungi la riga al DataFrame
                    new_row = pd.DataFrame({"Sede tecnica": [line]})
                    fl_dictionary['Mask_gen'] = pd.concat([fl_dictionary['Mask_gen'], new_row], ignore_index=True)
                elif ('*' in line) and (re.match(patterns['Mask_star'], line)):
                    # aggiungi una nuova chiave al df
                    fl_dictionary[line] = pd.DataFrame()
                else:
                    error_msg = (f"Errore riga {i}: la FL: {line} non rispetta la maschera.\n")
                    fl_errors += error_msg               

            except Exception as e:
                self.log(f"Errore nel processare la riga {i}: {str(e)}", 'error')
                return False, None
        # Se ci sono errori, mostra un messaggio di errore
        if fl_errors:
            self.log(f"Validazione fallita: {fl_errors}", 'error')
            return False, None
        else:
            self.log("Validazione dati completata con successo", 'success')
            if 'Mask_gen' in fl_dictionary:
                self.log(f"FL gen = {len(fl_dictionary['Mask_gen'])}", 'info')
                if len(fl_dictionary.keys()) > 1:
                    self.log(f"FL star = {len(fl_dictionary.keys()) -1}", 'info')
            else:
                self.log(f"FL star = {len(fl_dictionary.keys()) -1}", 'info')
            return True, fl_dictionary        

    # ----------------------------------------------------
    # Salvataggio dati parziali in caso di errore
    # ----------------------------------------------------
    def _save_partial_results(self, df_result):
        """
        Salva su file Excel i dati parzialmente elaborati prima di un'interruzione.
        Viene chiamato sia su success=False che nell'except esterno di update_data().
        Non solleva eccezioni: un fallimento del salvataggio viene solo loggato.
        """
        if df_result is None or df_result.empty:
            self.log("Nessun dato parziale da salvare", 'warning')
            return

        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_Excel = f"FL_parziale_{timestamp}.xlsx"
            elaborated = df_result[df_result["Result"] != ""].shape[0]
            total      = len(df_result)
            self.log(
                f"Salvataggio dati parziali ({elaborated}/{total} FL elaborate): {file_Excel}",
                'warning'
            )
            if self.save_excel_file_advanced(df_result, file_Excel,
                                             sheet_name='Dati_parziali',
                                             index=False,
                                             overwrite=True):
                self.log(f"Dati parziali salvati in: {file_Excel}", 'success')
            else:
                self.log("Impossibile salvare i dati parziali su file", 'error')
        except Exception as save_err:
            self.log(f"Errore durante il salvataggio dei dati parziali: {save_err}", 'error')

    # ----------------------------------------------------
    # Routine associata al tasto <Estrai Dati>
    # ----------------------------------------------------
    def update_data(self):

        # Disabilito il tasto
        self.extract_button.setEnabled(False)

        # Pre-inizializzato a None: se un'eccezione avviene prima dell'assegnazione
        # il gestore di errori può comunque verificare se ci sono dati parziali da salvare.
        df_result = None

        # ----------------------------------------------------
        # Validazione dati con maschere
        # ----------------------------------------------------
        if(True):
            # Prima verifica i dati nella finestra di testo sinistra (clipboard_area) che può contenere una lista di FL
            # oppure FL seguite dal carattere *
            # Crea un dizionario che ha come chiavi:
            # Mask_gen - contiene i valori delle lista in cui non compare il carattere *
            # FL con * - contiene un df vuoto che verrà popolato con le FL estratte con IH06
            result, self.fl_dictionary = self.validate_clipboard_data()
            if not result:
                self.log("Dati inseriti non validi", 'error')
                return
            # # Creo un dizionario che ha come chiavi i valori della lista data_string e come valori dei DataFrame vuoti
            # self.fl_dictionary = {item: pd.DataFrame() for item in data_string}


        # altrimenti estraggo i dati da SAP
        self.log("Avvio connessione SAP...")
        try:
            # Inizializza SOLO se non è già stato fatto
            if not self.session_manager:
                self.initialize_sap_components()


            with self.session_manager.get_session() as session:
                if session:
                    if session:
                        try:
                            self.infoUser = session.info.user
                            self.infoSystemName = session.info.systemName
                            self.infoClient = session.info.client
                            self.infoLanguage = session.info.language

                            self.log(f"ID utente:  {self.infoUser}", 'info')
                            self.log(f"System Name: {self.infoSystemName}", 'info')
                            self.log(f"Mandante: {self.infoClient}", 'info')
                            self.log(f"Lingua:  {self.infoLanguage}", 'info')
                        except Exception as e:
                            self.log(f"Errore lettura info SAP: {str(e)}", 'error')
                            return                        
                        self.log("Connessione SAP attiva", 'success')
                        extractor = SAPDataExtractor(session, self)
                        # Eseguo l'estrazione dei dati per ogni FL iterando per le chiavi del dizionario
                        if not self.fl_dictionary:
                            self.log("Nessuna FL da estrarre", 'warning')   
                            return
                        # Itero attraverso le chiavi del dizionario per ottenere tutte le liste di FL necessarie escludendo quelle che non sono in stato CRT
                        for key in self.fl_dictionary.keys():
                            
                            ### Estraggo tutte le FL che corrispondono all FL con * contenuta come chiave Utilizzo IH06
                            # Rimuovo le FL che non sono in stato CRT (in base alla lingua della sessione SAP)
                            if key != 'Mask_gen':
                                self.log("Estrazione dati FL contenenti *", 'loading')
                                success, df = extractor.extract_FL_list(key)
                            else:
                                self.log("Estrazione lista FL", 'loading')
                                stringa = '\r\n'.join(self.fl_dictionary[key]['Sede tecnica'].astype(str).str.strip()) # extract_FL_list deve ricevere come argomento una stringa
                                success, df = extractor.extract_FL_list(stringa)
                            if success:                                
                                # Modifico l'intestazione delle colonne del df mettendola in lingua IT
                                try:
                                    intestazione_df_IH06 = ['Sede tecnica']
                                    df_renamed = self.rename_columns_safely(df, intestazione_df_IH06)
                                    print(df_renamed.columns.tolist())
                                except ValueError as e:
                                    print(f"Errore: {e}")
                                    return
                                # Aggiungo i dati ottenuti al dizionario                               
                                self.fl_dictionary[key] = df_renamed
                                self.log(f"Estrazione FL {key} riuscita!", 'success')
                            else:
                                self.log(f"Errore durante l'estrazione della FL: {key}", 'error')
                                return False    
                        # ottenute le liste di FL, procedo con l'estrazione dei dati con la transazione IFLO
                        for key in self.fl_dictionary.keys():
                            self.log("Inizio estrazione dati lista FL", 'loading') 
                            
                            ### Estraggo i dati delle FL per ciascuna lista relativa ad una chiave
                            success, df = extractor.extract_FL_IFLO(self.fl_dictionary[key])
                            
                            if success:
                                self.log(f"Estratte {len(df)} FL per {key}", 'success')
                                # Concateno i dati estratti al df totale
                                if self.fl_df_tot.empty:
                                    self.fl_df_tot = df.copy()
                                else:
                                    self.fl_df_tot = pd.concat([self.fl_df_tot, df], ignore_index=True)
                            else:
                                self.log(f"Errore durante l'estrazione delle FL", 'error')
                                return

                        self.log("Estrazioni completata con successo", 'success')
                        self.log(f"Totale FL estratte = {len(self.fl_df_tot)}", 'success')

                        # Modifico l'intestazione delle colonne del df mettendola in lingua IT
                        try:
                            intestazione_df_IFLO = ['Sede tecnica', 'Definizione della sede tecnica', 'L', 'L_1', 'Tipologia', 'Componente', 'Sezione', 'Tipo ogg.', 'Prof.cat.']
                            df_renamed = self.rename_columns_safely(self.fl_df_tot, intestazione_df_IFLO)
                            print(df_renamed.columns.tolist())
                        except ValueError as e:
                            print(f"Errore: {e}")
                            return

                        # Creo il nome del file per salvare i dati
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        file_Excel = f"FL_estratte_" + timestamp + ".xlsx"
                        self.log(f"Salvo i dati in un file excel:\n     {file_Excel}", 'success')
                        # Salvo il DataFrame in un file Excel
                        if self.save_excel_file_advanced(df_renamed, file_Excel,
                                                        sheet_name='Dati_estratti',
                                                        index=False,
                                                        overwrite=True):
                            self.log("File Excel salvato con successo", 'success')
                        else:
                            self.log("Errore durante il salvataggio del file Excel", 'error')                            

                                
                        ### Verifico che il df  contenga fl con lingua attualmente in uso nella sessione di SAP
                        result, df_filtrato = self.Check_Lang(df_renamed, self.infoLanguage)
                        if result:
                                
                                ### Aggiorno i valori delle fl contenute nel df
                                success, df_result = extractor.update_FL_parallel(df_filtrato, self.session_manager, n_workers=4)

                                if success:
                                    # creo una statistica degli aggiornamenti eseguiti
                                    result_stat = self.analyze_result(df_result)   

                                    df_result = self.check_modifications_detailed(df_result)     

                                    # Creo il nome del file per salvare i dati
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    file_Excel = f"FL_aggiornate_" + timestamp + ".xlsx"
                                    self.log(f"Salvo i dati in un file excel:\n     {file_Excel}", 'success')
                                    # Salvo il DataFrame in un file Excel
                                    if self.save_excel_file_advanced(df_result, file_Excel,
                                                                    sheet_name='Dati_modificati',
                                                                    index=False,
                                                                    overwrite=True):
                                        self.log("File Excel salvato con successo", 'success')
                                    else:
                                        self.log("Errore durante il salvataggio del file Excel", 'error')
                                else:
                                    self.log("Aggiornamento interrotto — salvataggio dati parziali", 'warning')
                                    self._save_partial_results(df_result)
                        else:
                            self.log("Errore durante l'elaborazione del df", 'error')

                    self.log("Elaborazione terminata", 'success')

                else:
                    self.log("Connessione SAP NON attiva", 'error')
                    return
        except Exception as e:
            self.log(f"Estrazione dati SAP: Errore: {str(e)}", 'error')
            self._save_partial_results(df_result)
            return

        # ----------------------------------------------------
        # Verifica completata - ripristino il tasto di estrazione dei dati
        # ---------------------------------------------------- 
        self.extract_button.setEnabled(True)



    # ----------------------------------------------------
    # Modifica l'intestazione di un df
    # ---------------------------------------------------- 

    def rename_columns_safely(self, df, new_column_names, inplace=False):
        """
        Rinomina le colonne di un DataFrame con controlli di sicurezza.
        
        Args:
            df (pd.DataFrame): DataFrame da modificare
            new_column_names (list): Lista dei nuovi nomi delle colonne
            inplace (bool): Se True modifica il DataFrame originale, altrimenti crea una copia
        
        Returns:
            pd.DataFrame: DataFrame con colonne rinominate
            
        Raises:
            ValueError: Se il numero di colonne non corrisponde
            TypeError: Se new_column_names non è una lista
        """
        
        # Verifica che new_column_names sia una lista
        if not isinstance(new_column_names, (list, tuple)):
            raise TypeError(f"new_column_names deve essere una lista o tupla, ricevuto: {type(new_column_names)}")
        
        # Verifica che il numero di colonne corrisponda
        if len(df.columns) != len(new_column_names):
            raise ValueError(
                f"Numero di colonne non corrisponde!\n"
                f"  DataFrame ha {len(df.columns)} colonne: {list(df.columns)}\n"
                f"  Forniti {len(new_column_names)} nomi: {new_column_names}"
            )
        
        # Verifica duplicati nei nuovi nomi
        if len(new_column_names) != len(set(new_column_names)):
            duplicates = [name for name in new_column_names if new_column_names.count(name) > 1]
            raise ValueError(f"Nomi duplicati trovati nei nuovi nomi: {set(duplicates)}")
        
        # Verifica che tutti i nomi siano stringhe non vuote
        invalid_names = [name for name in new_column_names if not isinstance(name, str) or not name.strip()]
        if invalid_names:
            raise ValueError(f"Nomi di colonne non validi (devono essere stringhe non vuote): {invalid_names}")
        
        # Crea copia se richiesto
        working_df = df if inplace else df.copy()
        
        # Report delle modifiche
        print("📋 RINOMINAZIONE COLONNE:")
        print("  Vecchio nome → Nuovo nome")
        print("  " + "-" * 30)
        for old, new in zip(df.columns, new_column_names):
            print(f"  {old} → {new}")
        
        # Applica i nuovi nomi
        working_df.columns = new_column_names
        
        print(f"✅ Rinominazione completata per {len(new_column_names)} colonne")
        
        return working_df

    #-----------------------------------------------------------------------------
    # Genera una statistica dei risultati
    #-----------------------------------------------------------------------------
        
    def check_modifications_detailed(self, df):
        """
        Rileva e documenta le modifiche dei dati confrontando coppie di colonne correlate.
        """
        
        column_mapping = {
            'N_Tipologia': 'Tipologia',
            'N_Componente': 'Componente', 
            'N_Sezione': 'Sezione',
            'N_Tipo ogg.': 'Tipo ogg.',
            'N_Prof.cat.': 'Prof.cat.'
        }
        
        # Inizializza colonne
        df['Check'] = 0
        df['Modified_Fields'] = ''
        
        # Verifica esistenza colonna Result
        if 'Result' not in df.columns:
            print("⚠️ Colonna 'Result' non trovata")
            return df
        
        # Filtro per Result='S'
        mask_result_s = df['Result'].astype(str).str.contains('S', na=False)
        
        print(f"📊 Analisi: {len(df)} righe totali, {mask_result_s.sum()} con Result='S'")
        
        # Processa solo le righe con Result='S'
        for index in df[mask_result_s].index:
            row = df.loc[index]
            modified_fields = []
            
            for new_col, old_col in column_mapping.items():
                new_val = str(row[new_col]).strip() if pd.notna(row[new_col]) else ''
                old_val = str(row[old_col]).strip() if pd.notna(row[old_col]) else ''
                
                if new_val != old_val:
                    modified_fields.append(f"{old_col}: '{old_val}' → '{new_val}'")
            
            if modified_fields:
                df.at[index, 'Check'] = 1
                df.at[index, 'Modified_Fields'] = '; '.join(modified_fields)
            else:
                df.at[index, 'Modified_Fields'] = 'Nessuna modifica'
        
        # Per le righe che NON hanno Result='S', imposta messaggio specifico
        df.loc[~mask_result_s, 'Modified_Fields'] = 'Non elaborata (Result≠S)'
        
        return df

    #-----------------------------------------------------------------------------
    # Genera una statistica dei risultati
    #-----------------------------------------------------------------------------

    def analyze_result(self, df :pd.DataFrame) -> bool:
        """
        Analizza i caratteri nella colonna Result e calcola le percentuali
        """
        # Verifica che la colonna esista
        if "Result" not in df.columns:
            print("\n❌ Colonna 'Result' non trovata")
            return False
        
        # Conta tutti i caratteri (escludendo NaN)
        all_chars = df["Result"].dropna().astype(str)
        total_values = len(all_chars)
        
        if total_values == 0:
            print("\n⚠️ Nessun valore valido nella colonna Result")
            return False
        
        # Conta la frequenza di ogni carattere
        char_counts = all_chars.value_counts()
        
        print(f"\n📊 Analisi caratteri colonna 'Result' ({total_values} valori totali):")
        print("-" * 50)
        
        for char, count in char_counts.items():
            percentage = (count / total_values) * 100
            print(f"'{char}': {count:>4} occorrenze ({percentage:>5.1f}%)")
        
        return True        

    #-----------------------------------------------------------------------------
    # Filtra il df in base alla lingua indicata
    #-----------------------------------------------------------------------------

    def Check_Lang(self, df: pd.DataFrame, lang: str) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Filtra il DataFrame contiene dati nella lingua specificata
        
        Args:
            df (pd.DataFrame): DataFrame da verificare
            lang (str): Lingua da verificare
            
        Returns:
            bool: True se la lingua è presente, False altrimenti
            df_filtrato (pd.DataFrame): DataFrame filtrato con i soli valori appartenenti alla lingua indicata
        """
        
        self.log(f"✅ Lingua selezionata: {lang}", 'success')
                         
        try:
            if 'L_1' not in df.columns:
                raise KeyError("Colonna 'L_1' non presente")
            
            if df.empty:
                raise ValueError("DataFrame originale è vuoto")
            
            # Debug: mostra valori unici
            self.log(f"Valori lingua presenti: {df['L_1'].unique()}", 'info')
            print(f"🔍 Valori unici in L_1: {df['L_1'].unique()}")
            
            # Filtra usando il parametro lang (non hardcoded)
            df_filtrato = df[df['L_1'].str.upper() == lang.upper()]
            
            # Risultati
            if len(df_filtrato) == 0:
                self.log(f"Nessun valore per lingua = {lang}", 'error')
                print(f"❌ Nessun record con L_1 = {lang} trovato")
                raise ValueError(f"Nessun valore trovato per {lang}")
            else:
                self.log(f"Filtro completato. {len(df_filtrato)} elementi trovati", 'success')  # Fixed typo
                print(f"✅ Filtro completato: {len(df_filtrato)} elementi trovati")
                return True, df_filtrato
                
        except (KeyError, ValueError) as e:
            # Gestisci errori specifici
            self.log(f"Errore nella verifica lingua: {e}", 'error')
            print(f"❌ Errore: {e}")
        except Exception as e:
            # Gestisci errori imprevisti
            self.log(f"Errore imprevisto: {e}", 'error')
            print(f"❌ Errore imprevisto: {e}")
        
        return False, None

    def save_data(self):

        # Funzione per salvare i dati del df i un file excel
        pass

    def save_excel_file_advanced(self, df: pd.DataFrame, filename: str, 
                            sheet_name: str = 'Sheet1', 
                            index: bool = False,
                            overwrite: bool = True) -> bool:
        """
        Salva un DataFrame in un file Excel con opzioni avanzate
        
        Args:
            df (pd.DataFrame): DataFrame da salvare
            filename (str): Nome del file da creare/sovrascrivere
            sheet_name (str): Nome del foglio Excel (default: 'Sheet1')
            index (bool): Se includere l'indice come colonna (default: False)
            overwrite (bool): Se sovrascrivere file esistenti (default: True)
            
        Returns:
            bool: True se salvato con successo, False in caso di errore
        """
        file_path = os.path.join(self.current_dir, filename)
        file_path = Path(file_path) 
        
        try:
            # Verifica che il DataFrame non sia vuoto
            if df.empty:
                self.log(f"DataFrame vuoto.\nSalvataggio di {filename} non eseguito!", 'error')
                return False
            
            # Controlla se il file esiste già
            if file_path.exists() and not overwrite:
                self.log(f"File {filename} già esistente. \nSalvataggio non eseguito!", 'error')
                return False
            
            # Crea la directory se non esiste
            file_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Salva il DataFrame in Excel
            df.to_excel(
                file_path,
                sheet_name=sheet_name,
                index=index,
                na_rep='',
                header=True,
                engine='openpyxl'  # Engine specifico per .xlsx
            )
            
            return True
            
        except PermissionError:
            self.log(f"Permessi insufficienti per scrivere il file: {filename}", 'error')
            return False
            
        except FileNotFoundError:
            self.log(f"Percorso non trovato: {file_path.parent}", 'error')
            return False
            
        except Exception as e:
            self.log(f"Errore durante il salvataggio di {filename}: {str(e)}", 'error')
            return False
        

    #-----------------------------------------------------------------------------
    # Funzioni per il MultiThread
    #-----------------------------------------------------------------------------
    
    def initialize_sap_components(self):
        """Inizializza i componenti SAP"""
        try:
            self.log("Inizializzazione componenti SAP...", "info")
            
            # Crea session manager
            self.session_manager = SAPSessionManager(
                max_sessions=AppSettings.MAX_SAP_SESSIONS,
                connection_index=AppSettings.SAP_CONNECTION_INDEX,
                logger = self.logger
            )
            
            # Inizializza sessioni
            if not self.session_manager.initialize_sessions():
                raise Exception("Impossibile inizializzare le sessioni SAP")
            
            # Crea operations manager
            self.sap_operations = SAPOperations(logger=self.logger)
            
            # Crea thread manager
            self.thread_manager = ThreadManager(
                session_manager=self.session_manager,
                logger=self.logger
            )
            
            self.log("Componenti SAP inizializzati con successo", "success")
        
        except Exception as e:
            error_msg = str(e)
            self.log(f"Errore inizializzazione SAP: {error_msg}", "error")
            
            # Suggerimenti per l'utente
            self.log("", "info")
            self.log("Verifica che:", "warning")
            self.log("  1. SAP GUI sia aperto", "warning")
            self.log("  2. Sei loggato in almeno una sessione", "warning")
            self.log("  3. Lo scripting sia abilitato in SAP", "warning")
        
            raise
    
    def execute_sap_extraction_impianti(self)  -> tuple[bool, str|None]:
        """
        Estrae la lista degli impianti, per la tecnologia specificata, da SAP
        
        Returns:
            Tupla (successo, dati) dove:
            - successo: bool - True se operazione riuscita
            - dati: stringa - Stringa contenente lista degli impianti o None se errore
        """
        self.log("Avvio elaborazione...", "info")
        
        with self.session_manager.get_session() as session:
            if session:
                # STEP 1 Ricava lista FL
                # Ricavo lista degli impianti SingleThread
                str_impianti = self.sap_operations.ricava_lista_impianti(session, self.tecnologia)             
                if not str_impianti:
                    self.log("Impossibile ricavare la lista FL", "error")
                    return False, None
            else:
                self.log("Sessione SAP non disponibile", "error")
                return False, None
        
        self.log("Lista impianti ricavata con successo!", "success")
        return True, str_impianti


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()