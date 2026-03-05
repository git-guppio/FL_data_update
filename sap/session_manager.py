# File: sap/session_manager.py

import win32com.client
from typing import Optional, List, Dict
import threading
import time
import win32clipboard
import queue
from contextlib import contextmanager  # <-- IMPORTANTE: questo deve essere presente
import pythoncom  # FONDAMENTALE per COM threading
from core.base_component import BaseComponent

try:
    import keyboard
except ImportError:
    keyboard = None  # Se non disponibile, continua senza controllo ESC

class SAPSessionManager(BaseComponent):
    """
    SAP Session Manager corretto con supporto COM Threading e creazione sessioni
    """
    
    def __init__(self, max_sessions: int = 4, connection_index: int = 0, logger=None):
        """
        Inizializza il manager delle sessioni
        
        Args:
            max_sessions: Numero massimo di sessioni (default 6)
            connection_index: Indice della connessione SAP da utilizzare (default 0)
            logger: Istanza di ThreadSafeLogger (opzionale)
        """
        #  Inizializza BaseComponent con super()
        super().__init__(logger)

        self.max_sessions = max_sessions
        self.connection_index = connection_index
        self.lock = threading.Lock()
        self.initialized = False
        
        # Informazioni di sistema SAP (per thread safety)
        self.system_info = None

    def initialize_com_for_thread(self):
        """
        Inizializza COM per il thread corrente
        DEVE essere chiamato in ogni thread che usa SAP
        """
        try:
            pythoncom.CoInitialize()
            return True
        except Exception as e:
            print(f"ERRORE inizializzazione COM: {str(e)}")
            return False

    def cleanup_com_for_thread(self):
        """
        Pulisce COM per il thread corrente
        """
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"ERRORE cleanup COM: {str(e)}")

    def connect_to_sap(self) -> bool:
        """
        Stabilisce la connessione iniziale con SAP
        """
        try:
            print("Connessione a SAP GUI...")
            
            # Inizializza COM per il thread principale
            pythoncom.CoInitialize()
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not SapGuiAuto:
                print("ERRORE: Impossibile ottenere l'oggetto SAPGUI")
                return False

            application = SapGuiAuto.GetScriptingEngine
            if not application:
                print("ERRORE: Impossibile ottenere Scripting Engine")
                return False

            connection = application.Children(self.connection_index)
            if not connection:
                print(f"ERRORE: Impossibile ottenere la connessione {self.connection_index}")
                return False

            # Estrai informazioni di sistema per uso futuro
            session_count = connection.Children.Count
            if session_count > 0:
                first_session = connection.Children(0)
                try:
                    self.system_info = {
                        'system_name': first_session.Info.SystemName,
                        'client': first_session.Info.Client,
                        'connection_id': self.connection_index,
                        'current_sessions': session_count
                    }
                    print(f"Sistema SAP: {self.system_info['system_name']} Client: {self.system_info['client']}")
                except:
                    self.system_info = {
                        'connection_id': self.connection_index,
                        'current_sessions': session_count
                    }
            else:
                print("ERRORE: Nessuna sessione SAP disponibile")
                return False

            print("Connessione SAP stabilita con successo")
            return True

        except Exception as e:
            print(f"ERRORE durante la connessione a SAP: {str(e)}")
            return False
        finally:
            pythoncom.CoUninitialize()

    def get_current_session_count(self) -> int:
        """
        Restituisce il numero corrente di sessioni attive (thread-safe)
        """
        try:
            pythoncom.CoInitialize()
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(self.connection_index)
            
            count = connection.Children.Count
            pythoncom.CoUninitialize()
            return count
            
        except Exception as e:
            print(f"ERRORE nel conteggio delle sessioni: {str(e)}")
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return 0

    def get_working_session(self):
        """
        Ottiene una sessione utilizzabile per operazioni di setup (thread-safe)
        """
        try:
            pythoncom.CoInitialize()
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(self.connection_index)
            
            session_count = connection.Children.Count
            if session_count > 0:
                # Prova tutte le sessioni fino a trovarne una valida
                for i in range(session_count):
                    try:
                        session = connection.Children(i)
                        if session:
                            # Test di validità
                            _ = session.Info.SystemName
                            return session
                    except:
                        continue
            
            return None
            
        except Exception as e:
            print(f"ERRORE nell'ottenere sessione di lavoro: {str(e)}")
            return None

    def create_new_session(self, timeout: int = 10) -> bool:
        """
        Crea una nuova sessione SAP utilizzando il Session Manager (thread-safe)
        """
        try:
            current_count = self.get_current_session_count()
            
            if current_count >= self.max_sessions:
                print(f"AVVISO: Raggiunto il numero massimo di sessioni ({self.max_sessions})")
                return False

            print(f"Creazione nuova sessione (attualmente: {current_count}/{self.max_sessions})...")
            
            # Ottieni una sessione di lavoro
            working_session = self.get_working_session()
            if not working_session:
                print("ERRORE: Impossibile ottenere una sessione utilizzabile")
                return False

            # Crea nuova sessione tramite Session Manager
            try:
                print("Creazione tramite Session Manager...")
                working_session.findById("wnd[0]/tbar[0]/okcd").text = "/oSESSION_MANAGER"
                working_session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
            except Exception as e:
                print(f"ERRORE nell'esecuzione Session Manager: {str(e)}")
                return False
            finally:
                # Cleanup COM per questo thread
                pythoncom.CoUninitialize()
            
            # Attendi la creazione della nuova sessione
            start_time = time.time()
            print("Attesa creazione nuova sessione...")
            
            while True:
                elapsed_time = time.time() - start_time
                
                # Controllo timeout
                if elapsed_time > timeout:
                    print(f"ERRORE: Timeout ({timeout}s) nell'apertura della nuova sessione")
                    return False
                
                # Controllo tasto ESC (se disponibile)
                if keyboard:
                    try:
                        if keyboard.is_pressed('esc'):
                            print("Operazione annullata dall'utente (ESC)")
                            return False
                    except:
                        pass
                
                time.sleep(0.25)
                
                # Verifica se è stata creata una nuova sessione
                new_count = self.get_current_session_count()
                if new_count > current_count:
                    print(f"Nuova sessione creata con successo! Sessioni totali: {new_count}")
                    time.sleep(0.5)
                    return True

        except Exception as e:
            print(f"ERRORE durante la creazione della sessione: {str(e)}")
            return False

    def initialize_sessions(self, force_max: bool = True) -> bool:
        """
        Inizializza le sessioni fino al numero massimo
        """
        try:
            if not self.connect_to_sap():
                return False
            
            with self.lock:
                current_count = self.get_current_session_count()
                print(f"Sessioni attuali: {current_count}/{self.max_sessions}")
                
                if force_max and current_count < self.max_sessions:
                    # Crea sessioni fino al massimo
                    sessions_to_create = self.max_sessions - current_count
                    
                    print(f"Creazione di {sessions_to_create} nuove sessioni...")
                    
                    for i in range(sessions_to_create):
                        print(f"Creazione sessione {i+1}/{sessions_to_create}...")
                        
                        if not self.create_new_session():
                            print(f"Impossibile creare la sessione {i+1}")
                            break
                        
                        time.sleep(1)  # Pausa tra le creazioni
                
                else:
                    print("Numero desiderato di sessioni già disponibile")
                
                self.initialized = True
                final_count = self.get_current_session_count()
                print(f"Inizializzazione completata: {final_count} sessioni disponibili")
                
                return final_count > 0

        except Exception as e:
            print(f"ERRORE durante l'inizializzazione: {str(e)}")
            return False

    def create_thread_safe_session(self) -> Optional[object]:
        """
        Crea una sessione SAP specifica per il thread corrente con distribuzione migliorata
        """
        try:
            # Inizializza COM per questo thread
            if not self.initialize_com_for_thread():
                return None
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not SapGuiAuto:
                return None

            application = SapGuiAuto.GetScriptingEngine
            if not application:
                return None

            connection = application.Children(self.connection_index)
            if not connection:
                return None

            session_count = connection.Children.Count
            if session_count > 0:
                # Migliora la distribuzione delle sessioni
                thread_name = threading.current_thread().name
                thread_id = threading.current_thread().ident
                
                # Usa il nome del thread per una migliore distribuzione
                if "SAP_Worker" in thread_name:
                    # Estrai il numero del worker dal nome del thread
                    try:
                        worker_num = int(thread_name.split('_')[-1])
                        session_index = worker_num % session_count
                    except:
                        session_index = hash(thread_id) % session_count
                else:
                    session_index = hash(thread_id) % session_count
                
                # Prova diverse sessioni se quella calcolata non funziona
                for attempt in range(session_count):
                    try_index = (session_index + attempt) % session_count
                    try:
                        session = connection.Children(try_index)
                        if session:
                            # Test la sessione
                            _ = session.Info.SystemName
                            print(f"[Thread {thread_id}] Sessione {try_index + 1} acquisita (tentativo {attempt + 1})")
                            return session
                    except Exception as e:
                        print(f"[Thread {thread_id}] Sessione {try_index + 1} non utilizzabile: {str(e)}")
                        continue
            
            print(f"[Thread {threading.current_thread().ident}] Nessuna sessione disponibile")
            return None
            
        except Exception as e:
            print(f"[Thread {threading.current_thread().ident}] ERRORE creazione sessione: {str(e)}")
            return None

    @contextmanager
    def get_session(self, timeout: int = 30):
        """
        Context manager thread-safe per ottenere una sessione SAP
        """
        session = None
        try:
            if not self.initialized:
                print("Manager non inizializzato")
                yield None
                return

            # Ottieni una sessione specifica per questo thread
            session = self.create_thread_safe_session()
            
            if session:
                yield session
            else:
                print(f"[Thread {threading.current_thread().ident}] Sessione non disponibile")
                yield None
                
        except Exception as e:
            print(f"[Thread {threading.current_thread().ident}] ERRORE acquisizione sessione: {str(e)}")
            yield None
        finally:
            # Cleanup COM per questo thread
            if session:
                self.cleanup_com_for_thread()

    def get_status(self) -> Dict:
        """
        Restituisce lo stato del manager delle sessioni
        """
        try:
            current_count = self.get_current_session_count()
            
            return {
                'max_sessions': self.max_sessions,
                'total_sessions': current_count,
                'initialized': self.initialized,
                'system_info': self.system_info
            }
        except Exception as e:
            print(f"ERRORE nel recupero dello stato: {str(e)}")
            return {}
    
    def cleanup(self, close_sessions: bool = False) -> None:
        """
        Pulizia delle risorse
        
        Args:
            close_sessions: Se True, chiude anche le sessioni SAP aperte
        """
        try:
            with self.lock:
                # Opzionale: chiudi le sessioni SAP
                if close_sessions:
                    self._close_all_sessions()
                
                # Resetta stato interno
                self.initialized = False
                self.system_info = None
                
                print("Cleanup completato")
                
        except Exception as e:
            print(f"ERRORE durante il cleanup: {str(e)}")
    
    def _close_all_sessions(self):
        """Chiude tutte le sessioni SAP eccetto la prima"""
        try:
            pythoncom.CoInitialize()
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(self.connection_index)
            
            session_count = connection.Children.Count
            print(f"Chiusura di {session_count - 1} sessioni SAP...")
            
            # Chiudi tutte le sessioni tranne la prima (indice 0)
            for i in range(session_count - 1, 0, -1):  # Dal fondo verso l'alto
                try:
                    session = connection.Children(i)
                    session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
                    session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.3)
                    print(f"Sessione {i} chiusa")
                except Exception as e:
                    print(f"Errore chiusura sessione {i}: {str(e)}")
            
        except Exception as e:
            print(f"Errore durante chiusura sessioni: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def __enter__(self):
        """Context manager entry"""
        if not self.initialize_sessions():
            raise Exception("Impossibile inizializzare Session Manager")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.cleanup()