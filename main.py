import re
from pathlib import Path
import os
import sys
import pandas as pd
from datetime import datetime
from collections import deque
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                           QHBoxLayout, QWidget, QTextEdit, QListWidget, QLabel, QMessageBox,
                           QDialog, QRadioButton, QButtonGroup, QDialogButtonBox, QMenu, QAction,
                           QProgressBar, QSpinBox, QCheckBox, QGroupBox, QFormLayout)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QCursor, QFont, QTextCursor

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


# =============================================================================
# Dialog di configurazione (aperta dal gear button).
# Legge i valori correnti da AppSettings e li salva in-memory.
# =============================================================================
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Impostazioni")
        self.setMinimumWidth(380)
        self._build_ui()
        self._load_values()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)

        # --- Sezione SAP ---
        grp_sap = QGroupBox("SAP")
        form_sap = QFormLayout(grp_sap)

        self.spin_workers = QSpinBox()
        self.spin_workers.setRange(1, 8)
        self.spin_workers.setToolTip("Numero di sessioni SAP parallele (default 4)")
        form_sap.addRow("Worker paralleli:", self.spin_workers)

        self.spin_timeout = QSpinBox()
        self.spin_timeout.setRange(10, 300)
        self.spin_timeout.setSuffix(" s")
        self.spin_timeout.setToolTip("Timeout per singola FL in secondi (default 60)")
        form_sap.addRow("Timeout per FL:", self.spin_timeout)

        main_layout.addWidget(grp_sap)

        # --- Sezione pausa ---
        self.chk_pause = QCheckBox("Abilita pausa periodica anti-saturazione SAP")
        main_layout.addWidget(self.chk_pause)

        self.grp_pause = QGroupBox("Parametri pausa")
        form_pause = QFormLayout(self.grp_pause)

        self.spin_batch = QSpinBox()
        self.spin_batch.setRange(100, 10000)
        self.spin_batch.setSingleStep(100)
        self.spin_batch.setToolTip("Numero di FL elaborate prima di ogni pausa")
        form_pause.addRow("FL per batch:", self.spin_batch)

        self.spin_base_delay = QSpinBox()
        self.spin_base_delay.setRange(10, 3600)
        self.spin_base_delay.setSuffix(" s")
        self.spin_base_delay.setToolTip("Durata base della pausa in secondi")
        form_pause.addRow("Durata pausa base:", self.spin_base_delay)

        self.chk_progressive = QCheckBox("Incrementa pausa ad ogni batch")
        form_pause.addRow(self.chk_progressive)

        self.spin_increment = QSpinBox()
        self.spin_increment.setRange(5, 600)
        self.spin_increment.setSuffix(" s")
        self.spin_increment.setToolTip("Secondi aggiuntivi aggiunti ad ogni batch successivo")
        form_pause.addRow("Incremento per batch:", self.spin_increment)

        main_layout.addWidget(self.grp_pause)

        # Bottoni OK/Annulla
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._save_values)
        btn_box.rejected.connect(self.reject)
        main_layout.addWidget(btn_box)

        # Connessioni enable/disable
        self.chk_pause.toggled.connect(self.grp_pause.setEnabled)
        self.chk_progressive.toggled.connect(self.spin_increment.setEnabled)

    def _load_values(self):
        self.spin_workers.setValue(AppSettings.DEFAULT_WORKERS)
        self.spin_timeout.setValue(AppSettings.FL_TIMEOUT)
        self.chk_pause.setChecked(AppSettings.PAUSE_ENABLED)
        self.spin_batch.setValue(AppSettings.PAUSE_BATCH_SIZE)
        self.spin_base_delay.setValue(AppSettings.PAUSE_BASE_DELAY)
        self.chk_progressive.setChecked(AppSettings.PAUSE_PROGRESSIVE)
        self.spin_increment.setValue(AppSettings.PAUSE_INCREMENT)
        # Stato iniziale enable/disable
        self.grp_pause.setEnabled(AppSettings.PAUSE_ENABLED)
        self.spin_increment.setEnabled(AppSettings.PAUSE_PROGRESSIVE)

    def _save_values(self):
        AppSettings.DEFAULT_WORKERS   = self.spin_workers.value()
        AppSettings.FL_TIMEOUT        = self.spin_timeout.value()
        AppSettings.PAUSE_ENABLED     = self.chk_pause.isChecked()
        AppSettings.PAUSE_BATCH_SIZE  = self.spin_batch.value()
        AppSettings.PAUSE_BASE_DELAY  = self.spin_base_delay.value()
        AppSettings.PAUSE_PROGRESSIVE = self.chk_progressive.isChecked()
        AppSettings.PAUSE_INCREMENT   = self.spin_increment.value()
        self.accept()


# =============================================================================
# Worker thread: esegue tutte le operazioni SAP in background.
# Il thread GUI resta libero → QTimer a 100ms processa i log in real-time.
# Comunicazione GUI ↔ Worker solo tramite segnali Qt (thread-safe).
# =============================================================================
class SAPWorker(QThread):
    # (FL completate, FL totali)
    progress_updated = pyqtSignal(int, int)
    # True = tutto OK, False = terminato con errori/parziale
    operation_done   = pyqtSignal(bool)

    def __init__(self, main_window):
        super().__init__()
        self.mw = main_window

    def run(self):
        """Logica SAP spostata fuori dal thread GUI."""
        mw = self.mw
        df_result = None
        success   = False

        try:
            if not mw.session_manager:
                mw.initialize_sap_components()

            with mw.session_manager.get_session() as session:
                if not session:
                    mw.log("Connessione SAP NON attiva", 'error')
                    self.operation_done.emit(False)
                    return

                try:
                    mw.infoUser       = session.info.user
                    mw.infoSystemName = session.info.systemName
                    mw.infoClient     = session.info.client
                    mw.infoLanguage   = session.info.language
                    mw.log(f"ID utente:    {mw.infoUser}", 'info')
                    mw.log(f"System Name: {mw.infoSystemName}", 'info')
                    mw.log(f"Mandante:    {mw.infoClient}", 'info')
                    mw.log(f"Lingua:      {mw.infoLanguage}", 'info')
                except Exception as e:
                    mw.log(f"Errore lettura info SAP: {e}", 'error')
                    self.operation_done.emit(False)
                    return

                mw.log("Connessione SAP attiva", 'success')
                extractor = SAPDataExtractor(session, mw)

                if not mw.fl_dictionary:
                    mw.log("Nessuna FL da estrarre", 'warning')
                    self.operation_done.emit(False)
                    return

                # --- Fase 1: estrazione liste FL (IH06) ---
                for key in mw.fl_dictionary.keys():
                    if key != 'Mask_gen':
                        mw.log("Estrazione dati FL contenenti *", 'loading')
                        ok, df = extractor.extract_FL_list(key)
                    else:
                        mw.log("Estrazione lista FL", 'loading')
                        stringa = '\r\n'.join(
                            mw.fl_dictionary[key]['Sede tecnica'].astype(str).str.strip()
                        )
                        ok, df = extractor.extract_FL_list(stringa)

                    if ok:
                        try:
                            df_renamed = mw.rename_columns_safely(df, ['Sede tecnica'])
                        except ValueError as e:
                            mw.log(f"Errore rinomina colonne: {e}", 'error')
                            self.operation_done.emit(False)
                            return
                        mw.fl_dictionary[key] = df_renamed
                        mw.log(f"Estrazione FL {key} riuscita!", 'success')
                    else:
                        mw.log(f"Errore durante l'estrazione della FL: {key}", 'error')
                        self.operation_done.emit(False)
                        return

                # --- Fase 2: estrazione dettagli FL (SE16/IFLO) ---
                fl_df_tot = pd.DataFrame()
                for key in mw.fl_dictionary.keys():
                    mw.log("Inizio estrazione dati lista FL", 'loading')
                    ok, df = extractor.extract_FL_IFLO(mw.fl_dictionary[key])
                    if ok:
                        mw.log(f"Estratte {len(df)} FL per {key}", 'success')
                        fl_df_tot = df.copy() if fl_df_tot.empty else pd.concat(
                            [fl_df_tot, df], ignore_index=True
                        )
                    else:
                        mw.log("Errore durante l'estrazione delle FL", 'error')
                        self.operation_done.emit(False)
                        return

                mw.log("Estrazione completata con successo", 'success')
                mw.log(f"Totale FL estratte = {len(fl_df_tot)}", 'success')
                mw.fl_df_tot = fl_df_tot

                try:
                    intestazione = ['Sede tecnica', 'Definizione della sede tecnica',
                                    'L', 'L_1', 'Tipologia', 'Componente',
                                    'Sezione', 'Tipo ogg.', 'Prof.cat.']
                    df_renamed = mw.rename_columns_safely(fl_df_tot, intestazione)
                except ValueError as e:
                    mw.log(f"Errore rinomina colonne: {e}", 'error')
                    self.operation_done.emit(False)
                    return

                timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_Excel = f"FL_estratte_{timestamp}.xlsx"
                mw.log(f"Salvo i dati in un file excel:\n     {file_Excel}", 'success')
                if mw.save_excel_file_advanced(df_renamed, file_Excel,
                                               sheet_name='Dati_estratti',
                                               index=False, overwrite=True):
                    mw.log("File Excel salvato con successo", 'success')
                else:
                    mw.log("Errore durante il salvataggio del file Excel", 'error')

                # --- Fase 3: aggiornamento FL (IL02) ---
                result, df_filtrato = mw.Check_Lang(df_renamed, mw.infoLanguage)
                if not result:
                    mw.log("Errore durante l'elaborazione del df", 'error')
                    self.operation_done.emit(False)
                    return

                total = len(df_filtrato)
                self.progress_updated.emit(0, total)

                ok, df_result = extractor.update_FL_parallel(
                    df_filtrato,
                    mw.session_manager,
                    n_workers=4,
                    progress_callback=lambda c, t: self.progress_updated.emit(c, t)
                )

                if ok:
                    mw.analyze_result(df_result)
                    df_result = mw.check_modifications_detailed(df_result)
                    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
                    file_Excel = f"FL_aggiornate_{timestamp}.xlsx"
                    mw.log(f"Salvo i dati in un file excel:\n     {file_Excel}", 'success')
                    if mw.save_excel_file_advanced(df_result, file_Excel,
                                                   sheet_name='Dati_modificati',
                                                   index=False, overwrite=True):
                        mw.log("File Excel salvato con successo", 'success')
                    else:
                        mw.log("Errore durante il salvataggio del file Excel", 'error')
                    success = True
                else:
                    mw.log("Aggiornamento interrotto — salvataggio dati parziali", 'warning')
                    mw._save_partial_results(df_result)

                mw.log("Elaborazione terminata", 'success')

        except Exception as e:
            mw.log(f"Errore SAP: {e}", 'error')
            mw._save_partial_results(df_result)

        self.operation_done.emit(success)


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
        self.thread_manager  = None
        self.sap_operations  = None
        self.sap_worker      = None

        # Variabili per il timer elapsed e ETA
        self.start_time       = None
        self._fl_timestamps   = deque(maxlen=50)
        self.elapsed_timer = QTimer()
        self.elapsed_timer.timeout.connect(self._update_elapsed)

        # Setup GUI
        self.init_ui()

        # Setup timer per processare log dalla coda (100ms → real-time log)
        self.log_timer = QTimer()
        self.log_timer.timeout.connect(self.process_log_queue)
        self.log_timer.start(100)


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
        
        # Bottone Impostazioni (gear)
        self.settings_button = QPushButton("⚙")
        self.settings_button.setFixedWidth(32)
        self.settings_button.setToolTip("Impostazioni")
        self.settings_button.clicked.connect(self.open_settings)
        button_layout.addWidget(self.settings_button)

        # Aggiungi il layout dei bottoni al layout principale
        main_layout.addLayout(button_layout)

        # ----------------------------------------------------
        # Status bar: stato | barra progresso | elapsed time
        # ----------------------------------------------------
        sb = self.statusBar()

        self.status_label = QLabel("Pronto")
        sb.addWidget(self.status_label, 1)          # stretch=1: occupa lo spazio rimasto

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedWidth(220)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)
        sb.addPermanentWidget(self.progress_bar)

        self.elapsed_label = QLabel("00:00:00")
        self.elapsed_label.setFixedWidth(65)
        sb.addPermanentWidget(self.elapsed_label)

        sep = QLabel("|")
        sep.setFixedWidth(10)
        sep.setAlignment(Qt.AlignCenter)
        sb.addPermanentWidget(sep)

        self.eta_label = QLabel("ETA --:--:--")
        self.eta_label.setFixedWidth(90)
        sb.addPermanentWidget(self.eta_label)

        sb.setStyleSheet("QStatusBar { border-top: 1px solid #b0b0b0; }")

    # ----------------------------------------------------
    # Elapsed time
    # ----------------------------------------------------
    def _update_elapsed(self):
        if self.start_time:
            secs = int((datetime.now() - self.start_time).total_seconds())
            h, r = divmod(secs, 3600)
            m, s = divmod(r, 60)
            self.elapsed_label.setText(f"{h:02d}:{m:02d}:{s:02d}")

    # ----------------------------------------------------
    # Slot: aggiornamento progresso dal worker
    # ----------------------------------------------------
    def _on_progress_updated(self, completed: int, total: int):
        if total > 0:
            self.progress_bar.setRange(0, total)
            self.progress_bar.setValue(completed)
            self.status_label.setText(f"Aggiornamento FL: {completed} / {total}")

            # --- ETA con media mobile (ultime 50 FL) ---
            self._fl_timestamps.append(datetime.now())
            min_sample = 10
            if len(self._fl_timestamps) >= min_sample and completed < total:
                window_sec = (
                    self._fl_timestamps[-1] - self._fl_timestamps[0]
                ).total_seconds()
                if window_sec > 0:
                    rate = (len(self._fl_timestamps) - 1) / window_sec  # FL/s
                    eta_sec = int((total - completed) / rate)
                    h, r = divmod(eta_sec, 3600)
                    m, s = divmod(r, 60)
                    self.eta_label.setText(f"ETA {h:02d}:{m:02d}:{s:02d}")
            elif completed >= total:
                self.eta_label.setText("ETA 00:00:00")

    # ----------------------------------------------------
    # Slot: worker terminato
    # ----------------------------------------------------
    def _on_worker_finished(self, success: bool):
        self.elapsed_timer.stop()
        self.progress_bar.setVisible(False)
        self.eta_label.setText("ETA --:--:--")
        self.extract_button.setEnabled(True)
        if success:
            self.status_label.setText("Elaborazione completata")
        else:
            self.status_label.setText("Terminato con errori — verificare il log")

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
    def open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec_()

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
        result, self.fl_dictionary = self.validate_clipboard_data()
        if not result:
            self.log("Dati inseriti non validi", 'error')
            return

        self.extract_button.setEnabled(False)
        self.status_label.setText("Avvio elaborazione SAP...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.elapsed_label.setText("00:00:00")
        self.eta_label.setText("ETA --:--:--")
        self._fl_timestamps = deque(maxlen=50)
        self.start_time = datetime.now()
        self.elapsed_timer.start(1000)

        self.sap_worker = SAPWorker(self)
        self.sap_worker.progress_updated.connect(self._on_progress_updated)
        self.sap_worker.operation_done.connect(self._on_worker_finished)
        self.sap_worker.start()



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