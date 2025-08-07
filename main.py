import re
from pathlib import Path
import os
import sys
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                           QHBoxLayout, QWidget, QTextEdit, QListWidget, QLabel, QMessageBox,
                           QDialog, QRadioButton, QButtonGroup, QDialogButtonBox, QListWidgetItem, QStyle, QMenu, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
import SAP_Connection
import SAP_Transactions
from typing import Tuple, Optional, Dict

import logging

# Configurazione base del logging per tutta l'applicazione
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

# Logger specifico per questo modulo
logger = logging.getLogger("main").setLevel(logging.DEBUG)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Inizializza l'interfaccia utente
        self.setWindowTitle("Aggiorna valori FL")
        self.setGeometry(100, 100, 1000, 600)
        self.init_ui()
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

    def init_ui(self):
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
        
        self.log_list = QListWidget()

        # Imposta altezza uniforme per tutti gli elementi
        self.log_list.setUniformItemSizes(True)
        
        # Imposta spaziatura tra gli elementi
        self.log_list.setSpacing(2)  # 2 pixel di spazio tra le righe
        
        # Imposta font più leggibile (opzionale)
        font = self.log_list.font()
        font.setPointSize(9)  # Aumenta dimensione font
        self.log_list.setFont(font)


        right_panel.addWidget(self.log_list)

        # Attiva il menu contestuale per il widget dei log
        self.log_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log_list.customContextMenuRequested.connect(self.show_context_menu)        
        
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
    
    # ----------------------------------------------------
    # Funzioni per mostrare un menu contestuale x copiare i dati
    # ----------------------------------------------------
    def show_context_menu(self, position):
        # Crea menu contestuale
        context_menu = QMenu()
        
        # Aggiungi l'azione "Copia"
        copy_action = QAction("Copia elemento", self)
        copy_action.triggered.connect(self.copy_selected_items)
        context_menu.addAction(copy_action)
        
        # Aggiungi l'azione "Copia tutto"
        copy_all_action = QAction("Copia tutto", self)
        copy_all_action.triggered.connect(self.copy_all_items)
        context_menu.addAction(copy_all_action)
        
        # Mostra il menu contestuale alla posizione corrente del cursore
        context_menu.exec_(QCursor.pos())

    def copy_selected_items(self):
        # Copia solo gli elementi selezionati
        selected_items = self.log_list.selectedItems()
        if selected_items:
            text = "\n".join(item.text() for item in selected_items)
            QApplication.clipboard().setText(text)
            print("Elementi selezionati copiati negli appunti")        

    def copy_all_items(self):
        # Copia tutti gli elementi
        all_items = []
        for i in range(self.log_list.count()):
            all_items.append(self.log_list.item(i).text())
        
        text = "\n".join(all_items)
        QApplication.clipboard().setText(text)
        print("Tutti gli elementi copiati negli appunti")        


    def log_message(self, message, icon_type='info'):
        """
        Aggiunge un messaggio al log senza icone
        """
        item = QListWidgetItem(message)
        self.log_list.addItem(item)
        self.log_list.scrollToBottom()

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
        
    #     self.log_list.addItem(item)
    #     self.log_list.scrollToBottom()


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
            self.log_list.addItem(f"{icon} {message}")
            self.log_list.scrollToBottom()
    """    

    def clear_windows(self):
        self.clipboard_area.clear()
        self.log_list.clear()
        self.extract_button.setEnabled(True)
        self.upload_button.setEnabled(False)
        self.log_message("Finestre pulite")

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
            # Le roghe possono contenere codici di sedi tecniche complete oppure dei codici contenenti il carattere '*'
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
                    fl_dictionary[line] = pd.DataFrame
                else:
                    error_msg = (f"Errore riga {i}: la FL: {line} non rispetta la maschera.\n")
                    fl_errors += error_msg               

            except Exception as e:
                self.log_message(f"Errore nel processare la riga {i}: {str(e)}", 'error')
                return False, None
        # Se ci sono errori, mostra un messaggio di errore
        if fl_errors:
            self.log_message(f"Validazione fallita: {fl_errors}", 'error')
            return False, None
        else:
            self.log_message("Validazione dati completata con successo", 'success')
            if 'Mask_gen' in fl_dictionary:
                self.log_message(f"FL gen = {len(fl_dictionary['Mask_gen'])}", 'info')
                if len(fl_dictionary.keys()) > 1:
                    self.log_message(f"FL star = {len(fl_dictionary.keys()) -1}", 'info')
            else:
                self.log_message(f"FL star = {len(fl_dictionary.keys()) -1}", 'info')
            return True, fl_dictionary        

    # ----------------------------------------------------
    # Routine associata al tasto <Estrai Dati>
    # ----------------------------------------------------
    def update_data(self):

        # Disabilito il tasto
        self.extract_button.setEnabled(False)

        # ----------------------------------------------------
        # Validazione dati con maschere
        # ----------------------------------------------------        
        if(True):
            # Prima verifica i dati nella finestra di testo sinistra (clipboard_area)
            result, self.fl_dictionary = self.validate_clipboard_data()
            if not result:
                self.log_message("Dati inseriti non validi", 'error')
                return
            # # Creo un dizionario che ha come chiavi i valori della lista data_string e come valori dei DataFrame vuoti
            # self.fl_dictionary = {item: pd.DataFrame() for item in data_string}


        # altrimenti estraggo i dati da SAP
        self.log_message("Avvio connessione SAP...")
        try:
            with SAP_Connection.SAPGuiConnection() as sap:
                if sap.is_connected():
                    session = sap.get_session()
                    if session:
                        try:
                            self.infoUser = session.info.user
                            self.infoSystemName = session.info.systemName
                            self.infoClient = session.info.client
                            self.infoLanguage = session.info.language

                            self.log_message(f"ID utente:  {self.infoUser}", 'info')
                            self.log_message(f"System Name: {self.infoSystemName}", 'info')
                            self.log_message(f"Mandante: {self.infoClient}", 'info')
                            self.log_message(f"Lingua:  {self.infoLanguage}", 'info')
                        except Exception as e:
                            self.log_message(f"Errore lettura info SAP: {str(e)}", 'error')
                            return                        
                        self.log_message("Connessione SAP attiva", 'success')
                        # Eseguo l'estrazione dei dati                        
                        extractor = SAP_Transactions.SAPDataExtractor(session, self)
                        # Eseguo l'estrazione dei dati per ogni FL iterando per le chiavi del dizionario
                        if not self.fl_dictionary:
                            self.log_message("Nessuna FL da estrarre", 'warning')   
                            return
                        # Itero attraverso le chiavi del dizionario per ottenere tutte le liste di FL necessarie
                        for key in self.fl_dictionary.keys():
                            # Se la chiave è diversa da 'Mask_gen' allora si tratta di Fl che contengono il carattere '*'
                            if key != 'Mask_gen':
                            # Esamino i valori di FL contenuti nel dizionario
                                self.log_message("Inizio estrazione dati FL contenenti *", 'loading')
                                
                                ### Estraggo tutte le FL che corrispondono all FL con * contenuta come chiave
                                success, df = extractor.extract_FL_list(key)
                                # Modifico l'intestazione delle colonne del df mettendola in lingua IT
                                try:
                                    intestazione_df_IH06 = ['Sede tecnica']
                                    df_renamed = self.rename_columns_safely(df, intestazione_df_IH06)
                                    print(df_renamed.columns.tolist())
                                except ValueError as e:
                                    print(f"Errore: {e}")
                                    return
                                
                                if success:
                                    # Aggiungo i dati ottenuti al dizionario con chiave 'Mask_star'                                    
                                    self.fl_dictionary[key] = df_renamed
                                    self.log_message(f"Estrazione FL {key} riuscita!", 'success')
                                else:
                                    self.log_message(f"Errore durante l'estrazione della FL: {key}", 'error')
                                    return False
                        # ottenute le liste di FL, procedo con l'estrazione dei dati
                        for key in self.fl_dictionary.keys():
                            self.log_message("Inizio estrazione dati lista FL", 'loading') 
                            
                            ### Estraggo i dati delle FL per ciascuna lista relativa ad una chiave
                            success, df = extractor.extract_FL_IFLO(self.fl_dictionary[key])
                            
                            if success:
                                self.log_message(f"Estratte {len(df)} FL per {key}", 'success')
                                if self.fl_df_tot.empty:
                                    self.fl_df_tot = df.copy()
                                else:
                                    self.fl_df_tot = pd.concat([self.fl_df_tot, df], ignore_index=True)
                            else:
                                self.log_message(f"Errore durante l'estrazione delle FL", 'error')
                                return

                        self.log_message("Estrazioni completata con successo", 'success')
                        self.log_message(f"Totale FL estratte = {len(self.fl_df_tot)}", 'success')

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
                        self.log_message(f"Salvo i dati in un file excel:\n     {file_Excel}", 'success')
                        # Salvo il DataFrame in un file Excel
                        if self.save_excel_file_advanced(df_renamed, file_Excel,
                                                        sheet_name='Dati_estratti',
                                                        index=False,
                                                        overwrite=True):
                            self.log_message("File Excel salvato con successo", 'success')
                        else:
                            self.log_message("Errore durante il salvataggio del file Excel", 'error')                            

                                
                        ### Verifico che il df  contenga fl con lingua attualmente in uso nella sessione di SAP
                        result, df_filtrato = self.Check_Lang(df_renamed, self.infoLanguage)
                        if result:
                                
                                ### Aggiorno i valori delle fl contenute nel df
                                success, df_result = extractor.update_FL(df_filtrato)

                                if success:
                                    # creo una statistica degli aggiornamenti eseguiti
                                    result_stat = self.analyze_result(df_result)   

                                    df_result = self.check_modifications_detailed(df_result)     

                                    # Creo il nome del file per salvare i dati
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    file_Excel = f"FL_aggiornate_" + timestamp + ".xlsx"
                                    self.log_message(f"Salvo i dati in un file excel:\n     {file_Excel}", 'success')
                                    # Salvo il DataFrame in un file Excel
                                    if self.save_excel_file_advanced(df_result, file_Excel,
                                                                    sheet_name='Dati_modificati',
                                                                    index=False,
                                                                    overwrite=True):
                                        self.log_message("File Excel salvato con successo", 'success')
                                    else:
                                        self.log_message("Errore durante il salvataggio del file Excel", 'error')
                                else:
                                    self.log_message("Errore durante l'aggiornamento delle fl", 'error')                           
                        else:
                            self.log_message("Errore durante l'elaborazione del df", 'error')

                    self.log_message("Elaborazione terminata", 'success')

                else:
                    self.log_message("Connessione SAP NON attiva", 'error')
                    return
        except Exception as e:
            self.log_message(f"Estrazione dati SAP: Errore: {str(e)}", 'error')
            return    

        # ----------------------------------------------------
        # Verifica completata - ripristino il tasto di estrazione dei dati
        # ---------------------------------------------------- 
        self.extract_button.setEnabled(True)



    # ----------------------------------------------------
    # Modifica l' intestazione di un df
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
        
        self.log_message(f"✅ Lingua selezionata: {lang}", 'success')
                         
        try:
            if 'L_1' not in df.columns:
                raise KeyError("Colonna 'L_1' non presente")
            
            if df.empty:
                raise ValueError("DataFrame originale è vuoto")
            
            # Debug: mostra valori unici
            self.log_message(f"Valori lingua presenti: {df['L_1'].unique()}", 'info')
            print(f"🔍 Valori unici in L_1: {df['L_1'].unique()}")
            
            # Filtra usando il parametro lang (non hardcoded)
            df_filtrato = df[df['L_1'].str.upper() == lang.upper()]
            
            # Risultati
            if len(df_filtrato) == 0:
                self.log_message(f"Nessun valore per lingua = {lang}", 'error')
                print(f"❌ Nessun record con L_1 = {lang} trovato")
                raise ValueError(f"Nessun valore trovato per {lang}")
            else:
                self.log_message(f"Filtro completato. {len(df_filtrato)} elementi trovati", 'success')  # Fixed typo
                print(f"✅ Filtro completato: {len(df_filtrato)} elementi trovati")
                return True, df_filtrato
                
        except (KeyError, ValueError) as e:
            # Gestisci errori specifici
            self.log_message(f"Errore nella verifica lingua: {e}", 'error')
            print(f"❌ Errore: {e}")
        except Exception as e:
            # Gestisci errori imprevisti
            self.log_message(f"Errore imprevisto: {e}", 'error')
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
                self.log_message(f"DataFrame vuoto.\nSalvataggio di {filename} non eseguito!", 'error')
                return False
            
            # Controlla se il file esiste già
            if file_path.exists() and not overwrite:
                self.log_message(f"File {filename} già esistente. \nSalvataggio non eseguito!", 'error')
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
            self.log_message(f"Permessi insufficienti per scrivere il file: {filename}", 'error')
            return False
            
        except FileNotFoundError:
            self.log_message(f"Percorso non trovato: {file_path.parent}", 'error')
            return False
            
        except Exception as e:
            self.log_message(f"Errore durante il salvataggio di {filename}: {str(e)}", 'error')
            return False

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()