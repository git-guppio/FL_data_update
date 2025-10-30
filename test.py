import win32com.client
from typing import Optional, List, Dict
import threading
import time
import queue
from contextlib import contextmanager
import keyboard

class SAPSessionManager:
    """
    Classe per gestire dinamicamente le sessioni SAP
    """
    
    def __init__(self, max_sessions: int = 6, connection_index: int = 0):
        """
        Inizializza il manager delle sessioni
        
        Args:
            max_sessions: Numero massimo di sessioni (default 6)
            connection_index: Indice della connessione SAP da utilizzare (default 0)
        """
        self.max_sessions = max_sessions
        self.connection_index = connection_index
        self.SapGuiAuto: Optional[object] = None
        self.application: Optional[object] = None
        self.connection: Optional[object] = None
        self.sessions: List[object] = []
        self.available_sessions = queue.Queue()
        self.lock = threading.Lock()
        self.initialized = False

    def connect_to_sap(self) -> bool:
        """
        Stabilisce la connessione iniziale con SAP
        
        Returns:
            bool: True se la connessione è stabilita con successo
        """
        try:
            print("Connessione a SAP GUI...")
            self.SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not self.SapGuiAuto:
                print("ERRORE: Impossibile ottenere l'oggetto SAPGUI")
                return False

            self.application = self.SapGuiAuto.GetScriptingEngine
            if not self.application:
                print("ERRORE: Impossibile ottenere Scripting Engine")
                return False

            self.connection = self.application.Children(self.connection_index)
            if not self.connection:
                print(f"ERRORE: Impossibile ottenere la connessione {self.connection_index}")
                return False

            print("Connessione SAP stabilita con successo")
            return True

        except Exception as e:
            print(f"ERRORE durante la connessione a SAP: {str(e)}")
            return False

    def get_current_session_count(self) -> int:
        """
        Restituisce il numero corrente di sessioni attive
        
        Returns:
            int: Numero di sessioni attive
        """
        try:
            if self.connection:
                return self.connection.Children.Count
            return 0
        except Exception as e:
            print(f"ERRORE nel conteggio delle sessioni: {str(e)}")
            return 0

    def get_active_session(self) -> Optional[object]:
        """
        Restituisce una sessione utilizzabile (prima disponibile)
        
        Returns:
            object: Sessione SAP utilizzabile o None
        """
        try:
            if not self.connection:
                return None
            
            # Prova a ottenere la prima sessione disponibile
            session_count = self.connection.Children.Count
            if session_count > 0:
                # Cerca una sessione valida
                for i in range(session_count):
                    try:
                        session = self.connection.Children(i)
                        if session:
                            # Verifica che la sessione sia utilizzabile
                            # Tenta di accedere a una proprietà base per validare
                            _ = session.Info.SystemName  # Test di validità
                            return session
                    except:
                        continue
            
            return None
            
        except Exception as e:
            print(f"ERRORE nell'ottenere una sessione utilizzabile: {str(e)}")
            return None

    def create_new_session(self, timeout: int = 10) -> bool:
        """
        Crea una nuova sessione SAP utilizzando il Session Manager
        
        Args:
            timeout: Timeout in secondi per la creazione della sessione
            
        Returns:
            bool: True se la sessione è stata creata con successo
        """
        try:
            current_count = self.get_current_session_count()
            
            if current_count >= self.max_sessions:
                print(f"AVVISO: Raggiunto il numero massimo di sessioni ({self.max_sessions})")
                return False

            print(f"Creazione nuova sessione (attualmente: {current_count}/{self.max_sessions})...")
            
            # Ottieni una sessione utilizzabile (prima disponibile)
            working_session = self.get_active_session()
            if not working_session:
                print("ERRORE: Impossibile ottenere una sessione utilizzabile")
                return False
            
            # # Metodo 1: Shortcut da tastiera (più affidabile)
            # try:
            #     print("Tentativo 1: Shortcut tastiera per nuova sessione...")
            #     working_session.findById("wnd[0]").sendVKey(16)  # Ctrl+N equivalente
            #     time.sleep(1)
            #     new_count = self.get_current_session_count()
            #     if new_count > current_count:
            #         session_created = True
            #         print("Sessione creata con shortcut tastiera!")
            # except Exception as e:
            #     print(f"Metodo 1 fallito: {str(e)}")
            
            # Metodo 2: Session Manager (come AutoHotkey)
            try:
                print("Tentativo creazione tramite Session Manager...")
                working_session.findById("wnd[0]/tbar[0]/okcd").text = "/oSESSION_MANAGER"
                working_session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
            except Exception as e:
                    print(f"MTentativo fallito: {str(e)}")
            
            # Attendi la creazione della nuova sessione con timeout
            start_time = time.time()
            print("Attesa creazione nuova sessione...")
            
            while True:
                elapsed_time = time.time() - start_time
                
                # Controllo timeout
                if elapsed_time > timeout:
                    print(f"ERRORE: Timeout ({timeout}s) nell'apertura della nuova sessione")
                    return False
                
                # Controllo tasto ESC (se supportato)
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
        Inizializza tutte le sessioni fino al numero massimo
        
        Args:
            force_max: Se True, crea sessioni fino al massimo consentito
            
        Returns:
            bool: True se l'inizializzazione è completata
        """
        try:
            if not self.connect_to_sap():
                return False
            
            with self.lock:
                current_count = self.get_current_session_count()
                print(f"Sessioni attuali: {current_count}/{self.max_sessions}")
                
                if force_max:
                    # Crea sessioni fino al massimo
                    sessions_to_create = self.max_sessions - current_count
                    
                    if sessions_to_create > 0:
                        print(f"Creazione di {sessions_to_create} nuove sessioni...")
                        
                        for i in range(sessions_to_create):
                            print(f"Creazione sessione {i+1}/{sessions_to_create}...")
                            
                            if not self.create_new_session():
                                print(f"Impossibile creare la sessione {i+1}")
                                break
                            
                            time.sleep(1)  # Pausa tra le creazioni
                    
                    else:
                        print("Numero massimo di sessioni già raggiunto")
                
                # Aggiorna la lista delle sessioni
                self.update_session_list()
                self.initialized = True
                
                final_count = len(self.sessions)
                print(f"Inizializzazione completata: {final_count} sessioni disponibili")
                
                return final_count > 0

        except Exception as e:
            print(f"ERRORE durante l'inizializzazione: {str(e)}")
            return False

    def update_session_list(self) -> None:
        """
        Aggiorna la lista delle sessioni disponibili
        """
        try:
            self.sessions.clear()
            
            # Svuota la coda
            while not self.available_sessions.empty():
                try:
                    self.available_sessions.get_nowait()
                except queue.Empty:
                    break
            
            if self.connection:
                session_count = self.connection.Children.Count
                
                for i in range(session_count):
                    try:
                        session = self.connection.Children(i)
                        if session:
                            self.sessions.append(session)
                            self.available_sessions.put(session)
                            print(f"Sessione {i+1} aggiunta al pool")
                    except Exception as e:
                        print(f"ERRORE nell'aggiungere la sessione {i}: {str(e)}")

        except Exception as e:
            print(f"ERRORE nell'aggiornamento delle sessioni: {str(e)}")

    @contextmanager
    def get_session(self, timeout: int = 30):
        """
        Context manager per ottenere una sessione dal pool
        
        Args:
            timeout: Timeout per ottenere una sessione disponibile
        """
        session = None
        try:
            if not self.initialized:
                print("Pool non inizializzato. Inizializzazione in corso...")
                if not self.initialize_sessions():
                    yield None
                    return
            
            # Ottieni una sessione disponibile
            session = self.available_sessions.get(timeout=timeout)
            session_index = self.sessions.index(session) + 1
            
            print(f"Sessione {session_index} acquisita")
            yield session
            
        except queue.Empty:
            print(f"TIMEOUT: Nessuna sessione disponibile entro {timeout} secondi")
            yield None
        except Exception as e:
            print(f"ERRORE nell'acquisizione della sessione: {str(e)}")
            yield None
        finally:
            # Rilascia la sessione
            if session:
                self.available_sessions.put(session)
                session_index = self.sessions.index(session) + 1 if session in self.sessions else "?"
                print(f"Sessione {session_index} rilasciata")

    def get_status(self) -> Dict:
        """
        Restituisce lo stato del manager delle sessioni
        
        Returns:
            Dict: Informazioni sullo stato corrente
        """
        try:
            current_count = self.get_current_session_count()
            available_count = self.available_sessions.qsize()
            busy_count = len(self.sessions) - available_count
            
            return {
                'max_sessions': self.max_sessions,
                'total_sessions': current_count,
                'managed_sessions': len(self.sessions),
                'available_sessions': available_count,
                'busy_sessions': busy_count,
                'initialized': self.initialized,
                'connection_active': self.connection is not None
            }
        except Exception as e:
            print(f"ERRORE nel recupero dello stato: {str(e)}")
            return {}

    def cleanup(self) -> None:
        """
        Pulizia delle risorse
        """
        try:
            with self.lock:
                # Svuota la coda
                while not self.available_sessions.empty():
                    try:
                        self.available_sessions.get_nowait()
                    except queue.Empty:
                        break
                
                self.sessions.clear()
                self.connection = None
                self.application = None
                self.SapGuiAuto = None
                self.initialized = False
                
                print("Cleanup completato")
                
        except Exception as e:
            print(f"ERRORE durante il cleanup: {str(e)}")

    def __enter__(self):
        """Support for context manager - inizializza le sessioni automaticamente"""
        if not self.initialize_sessions():
            raise Exception("Impossibile inizializzare il Session Manager")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Cleanup when exiting context manager"""
        self.cleanup()


# Funzione per operazioni SAP parallele
def execute_parallel_sap_operations(manager: SAPSessionManager, operations_list: List, max_workers: int = None):
    """
    Esegue operazioni SAP in parallelo utilizzando il session manager
    
    Args:
        manager: Session manager inizializzato
        operations_list: Lista di tuple (funzione_operazione, dati_operazione)
        max_workers: Numero massimo di worker paralleli
    """
    import concurrent.futures
    
    if max_workers is None:
        max_workers = min(len(operations_list), manager.max_sessions)
    
    print(f"Esecuzione {len(operations_list)} operazioni con {max_workers} worker paralleli")
    
    def execute_operation(operation_data):
        operation_func, data = operation_data
        with manager.get_session() as session:
            if session:
                try:
                    return operation_func(session, data)
                except Exception as e:
                    print(f"ERRORE nell'operazione: {str(e)}")
                    return None
            return None
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(execute_operation, op) for op in operations_list]
        results = []
        
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            if result:
                results.append(result)
                print(f"Operazione completata: {result}")
    
    return results


# Esempio di operazione SAP
def esempio_operazione_sap(session, data):
    """Esempio di operazione SAP"""
    try:
        print(f"Esecuzione operazione: {data}")
        
        # Esempio di operazioni SAP
        # session.findById("wnd[0]/tbar[0]/okcd").text = "MM01"
        # session.findById("wnd[0]").sendVKey(0)
        
        # Simula operazione
        time.sleep(2)
        
        return f"Completato: {data}"
        
    except Exception as e:
        print(f"ERRORE nell'operazione SAP: {str(e)}")
        return None


def main():
    """Esempio di utilizzo del Session Manager"""
    
    print("🔧 ESEMPI DI UTILIZZO OPERAZIONI SAP PARALLELE")
    print("=" * 60)
    
    try:
        # Esempio 1: Consultazione materiali
        print("\n1️⃣  ESEMPIO: CONSULTAZIONE MATERIALI")
        esempio_consultazione_ordini()
   
    except KeyboardInterrupt:
        print("\nOperazione interrotta dall'utente")
    except Exception as e:
        print(f"ERRORE generale: {str(e)}")

# ============================================================================
# ESEMPIO: Consultazione valori ordini di  manutenzione in SAP
# ============================================================================

def esempio_consultazione_ordini():
    """Esempio: Consultazione parallela di 20 materiali"""
    
    with SAPSessionManager(max_sessions=6) as manager:


        # Test singola operazione
        print("\n=== TEST OPERAZIONE SINGOLA ===")
        with manager.get_session() as session:
            if session:
                result = consulta_ordine_sap("240000523484")
                # Qui esegui le tue operazioni SAP
                time.sleep(1)
                print("Operazione completata")

        
        # Lista degli ordini di manutenzione da consultare
        ordini_manutenzione = [
            "240000523484", "240000523503", "240000523509", "210001280925", "210001280927",
            "210001280928", "210001280926", "210001280930", "210001280931", "240000523504",
            "240000523510", "240000523486", "240000523498", "240000523505", "240000523456",
            "240000523457", "240000523492", "240000523494", "240000523449", "240000523455"
        ]
        
        print(f"🚀 Avvio consultazione parallela di {len(ordini_manutenzione)} ordini")
        start_time = time.time()
        
        # Prepara le operazioni: lista di tuple (funzione, parametro)
        operations = [(consulta_ordine_sap, ordine) for ordine in ordini_manutenzione]
        
        # Esegui le operazioni in parallelo (max 6 contemporaneamente)
        results = execute_parallel_sap_operations(manager, operations, max_workers=6)
        
        end_time = time.time()
        
        # Analizza i risultati
        print(f"\n📊 RISULTATI CONSULTAZIONE MATERIALI:")
        print(f"   ⏱️  Tempo totale: {end_time - start_time:.2f} secondi")
        print(f"   📋 Operazioni richieste: {len(ordini_manutenzione)}")
        print(f"   ✅ Operazioni completate: {len(results)}")
        print(f"   ❌ Operazioni fallite: {len(ordini_manutenzione) - len(results)}")
        
        # Risultati dettagliati
        successi = [r for r in results if r and r.get('status') == 'success']
        errori = [r for r in results if r and r.get('status') == 'error']
        
        print(f"   🎯 Successi: {len(successi)}")
        print(f"   🚫 Errori: {len(errori)}")
        
        # Stampa dettagli successi
        print(f"\n✅ ORDINI CONSULTATI CON SUCCESSO:")
        for ordine_info in successi[:5]:  # Prime 5
            print(f"   • {ordine_info['ordine']}: {ordine_info.get('status', 'N/A')}")
        
        if len(successi) > 5:
            print(f"   ... e altri {len(successi) - 5} ordini")
        
        return results


def consulta_ordine_sap(session, ordine_code):
    """
    Consulta un ordine di manutenzione in SAP e restituisce alcune informazioni
    
    Args:
        session: Sessione SAP
        materiale_code: Codice ordine da consultare
        
    Returns:
        dict: Informazioni dell'ordine o None se errore
    """
    try:
        print(f"[Thread {threading.current_thread().ident}] Consultazione ordine: {ordine_code}")
        
        # Vai alla transazione MM03
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW33"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        
        # Inserisci il codice ordine
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = ordine_code
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        
        # Leggi i dati (esempio)
        try:
            stato = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT").text
            data_inizio = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GSTRP").text
            
            risultato = {
                'ordine': ordine_code,
                'status': stato,
                'data_inizio': data_inizio,
                'status': 'success'
            }
            
            print(f"[Thread {threading.current_thread().ident}] ✅ {ordine_code}: {stato}")
            return risultato
            
        except Exception as e:
            print(f"[Thread {threading.current_thread().ident}] ❌ Errore lettura dati {ordine_code}: {str(e)}")
            return {
                'materiale': ordine_code,
                'status': 'error',
                'errore': str(e)
            }
            
    except Exception as e:
        print(f"[Thread {threading.current_thread().ident}] ❌ Errore generale {ordine_code}: {str(e)}")
        return {
            'materiale': ordine_code,
            'status': 'error',
            'errore': str(e)
        }
    finally:
        # Torna al menu principale
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass

if __name__ == "__main__":
    main()