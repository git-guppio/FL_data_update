import win32com.client
from typing import Optional, List, Dict
import threading
import time
import queue
from contextlib import contextmanager
import pythoncom  # FONDAMENTALE per COM threading
try:
    import keyboard
except ImportError:
    keyboard = None  # Se non disponibile, continua senza controllo ESC

class SAPSessionManager:
    """
    SAP Session Manager corretto con supporto COM Threading e creazione sessioni
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

    def cleanup(self) -> None:
        """
        Pulizia delle risorse
        """
        try:
            with self.lock:
                self.initialized = False
                self.system_info = None
                print("Cleanup completato")
        except Exception as e:
            print(f"ERRORE durante il cleanup: {str(e)}")

    def __enter__(self):
        """Context manager entry"""
        if not self.initialize_sessions():
            raise Exception("Impossibile inizializzare Session Manager")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.cleanup()


def execute_parallel_sap_operations_corrected(manager: SAPSessionManager, operations_list: List, max_workers: int = None):
    """
    Esegue operazioni SAP in parallelo con gestione corretta COM threading
    """
    import concurrent.futures
    
    if max_workers is None:
        max_workers = min(len(operations_list), manager.max_sessions)
    
    print(f"🚀 Avvio {len(operations_list)} operazioni parallele con {max_workers} worker")
    
    def execute_operation_thread_safe(operation_data):
        """
        Wrapper thread-safe per eseguire operazioni SAP
        """
        operation_func, data = operation_data
        
        # Ogni thread deve avere la propria sessione
        with manager.get_session() as session:
            if session:
                try:
                    return operation_func(session, data)
                except Exception as e:
                    print(f"[Thread {threading.current_thread().ident}] ERRORE operazione: {str(e)}")
                    return {
                        'ordine': str(data),
                        'status': 'error',
                        'errore': str(e)
                    }
            else:
                print(f"[Thread {threading.current_thread().ident}] Sessione non disponibile")
                return {
                    'ordine': str(data),
                    'status': 'error',
                    'errore': 'Sessione non disponibile'
                }
    
    results = []
    
    # Usa ThreadPoolExecutor per gestire i thread
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="SAP_Worker") as executor:
        # Invia tutte le operazioni
        futures = [executor.submit(execute_operation_thread_safe, op) for op in operations_list]
        
        # Raccogli i risultati
        for i, future in enumerate(concurrent.futures.as_completed(futures)):
            try:
                result = future.result(timeout=60)  # Timeout per operazione
                if result:
                    results.append(result)
                    status = "✅" if result.get('status') == 'success' else "❌"
                    print(f"{status} Operazione {i+1}/{len(operations_list)}: {result.get('ordine', 'N/A')}")
                else:
                    print(f"❌ Operazione {i+1}/{len(operations_list)} fallita")
            except concurrent.futures.TimeoutError:
                print(f"⏱️  Operazione {i+1}/{len(operations_list)} timeout")
            except Exception as e:
                print(f"❌ Operazione {i+1}/{len(operations_list)} errore: {str(e)}")
    
    print(f"📊 Operazioni completate: {len(results)}/{len(operations_list)}")
    return results


def consulta_ordine_sap(session, ordine_code):
    """
    Consulta un ordine di manutenzione in SAP (thread-safe)
    """
    try:
        thread_id = threading.current_thread().ident
        print(f"[Thread {thread_id}] Consultazione ordine: {ordine_code}")
        
        # Vai alla transazione IW33
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW33"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        
        # Inserisci il codice ordine
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = ordine_code
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        
        # Leggi i dati dell'ordine
        try:
            stato = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT").text
            data_inizio = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GSTRP").text
            
            risultato = {
                'ordine': ordine_code,
                'stato': stato,
                'data_inizio': data_inizio,
                'thread_id': thread_id,
                'status': 'success'
            }
            
            print(f"[Thread {thread_id}] ✅ {ordine_code}: {stato}")
            return risultato
            
        except Exception as e:
            print(f"[Thread {thread_id}] ❌ Errore lettura dati {ordine_code}: {str(e)}")
            return {
                'ordine': ordine_code,
                'thread_id': thread_id,
                'status': 'error',
                'errore': f"Errore lettura: {str(e)}"
            }
            
    except Exception as e:
        thread_id = threading.current_thread().ident
        print(f"[Thread {thread_id}] ❌ Errore generale {ordine_code}: {str(e)}")
        return {
            'ordine': ordine_code,
            'thread_id': thread_id,
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


def esempio_consultazione_ordini(n_thread: int = 4):
    """
    Esempio corretto: Consultazione parallela ordini di manutenzione
    """
    try:
        # Usa meno sessioni inizialmente per testare la stabilità
        with SAPSessionManager(max_sessions=n_thread) as manager:
            
            # Lista degli ordini di manutenzione
            ordini_manutenzione = [
            "240000523484", "240000523503", "240000523509", "210001280925", "210001280927",
            "210001280928", "210001280926", "210001280930", "210001280931", "240000523504",
            "240000523510", "240000523486", "240000523498", "240000523505", "240000523456",
            "240000523457", "240000523492", "240000523494", "240000523449", "240000523455",
            "240000523467", "240000523468", "240000523469", "240000523470", "240000523471",
            "240000523489", "240000523495", "240000523497", "240000523499", "240000523501",
            "240000523506", "240000523563", "240000523573", "240000523575", "240000523586",
            "240000523588", "240000523605", "240000523608", "240000523450", "240000523491",
            "240000523524", "240000523530", "240000523618", "240000523459", "240000523461",
            "240000523463", "240000523464", "240000523465", "240000523466", "240000523534",
            "240000523538", "240000523574", "240000523587", "240000523612", "240000523620",
            "240000523623", "210001280929", "210001280932", "210001280933", "210001280934",
            "210001280935", "210001280936", "210001280937", "210001280938", "210001280939",
            "210001280940", "210001280941", "210001280942", "210001280943", "210001280944",
            "210001280945", "210001280946", "210001280947", "210001280948", "210001280949",
            "210001280950", "210001280951", "210001280952", "210001280953", "210001280954",
            "210001280955", "210001280956", "210001280957", "210001280958", "240000523481",
            "240000523526", "240000523527", "240000523562", "240000523566", "240000523569",
            "240000523597", "240000523602", "240000523604", "240000523606", "240000523610",
            "240000523670", "900000049177", "210001280336", "210001280337", "210001280338",
            "210001280393", "210001280394", "210001280395", "210001280396", "210001280397",
            "210001280398", "210001280399", "210001280400", "210001280401", "210001280402",
            "210001280403", "210001280404", "210001280405", "210001280406", "210001280407",
            "210001280408", "210001280409", "210001280410", "210001280411", "210001280412",
            "210001280413", "210001280414", "210001280484", "210001280485", "210001280486",
            "210001280487", "210001280488", "210001280489", "210001280490", "210001280491",
            "210001280492", "210001280493", "210001280494", "210001280495", "210001280496",
            "210001280497", "210001280498", "210001280499", "210001280659", "210001280660",
            "210001280661", "210001280662", "210001280663", "210001280664", "210001280665",
            "210001280666", "210001280667", "210001280668", "210001280669", "210001280670"
            ]
            
            print(f"🔧 Consultazione {len(ordini_manutenzione)} ordini di manutenzione")
            start_time = time.time()
            
            # Prepara le operazioni
            operations = [(consulta_ordine_sap, ordine) for ordine in ordini_manutenzione]
            
            # Esegui con la versione corretta
            results = execute_parallel_sap_operations_corrected(manager, operations, max_workers=n_thread)
            
            end_time = time.time()
            
            # Analizza i risultati
            print(f"\n📊 RISULTATI CONSULTAZIONE ORDINI - NUMERO DI THREAD: {n_thread}")
            print(f"   ⏱️  Tempo totale: {end_time - start_time:.2f} secondi")
            print(f"   📋 Operazioni richieste: {len(ordini_manutenzione)}")
            print(f"   ✅ Operazioni completate: {len(results)}")
            
            # Risultati dettagliati
            successi = [r for r in results if r and r.get('status') == 'success']
            errori = [r for r in results if r and r.get('status') == 'error']
            
            print(f"   🎯 Successi: {len(successi)}")
            print(f"   🚫 Errori: {len(errori)}")
            
            # Stampa dettagli successi
            print(f"\n✅ ORDINI CONSULTATI CON SUCCESSO:")
            for ordine_info in successi[:5]:  # Prime 5
                print(f"   • {ordine_info['ordine']}: {ordine_info.get('stato', 'N/A')}")
            
            if len(successi) > 5:
                print(f"   ... e altri {len(successi) - 5} ordini")
            
            # Stampa alcuni errori se presenti
            if errori:
                print(f"\n❌ ERRORI RISCONTRATI:")
                for error_info in errori[:3]:  # Prime 3
                    print(f"   • {error_info['ordine']}: {error_info.get('errore', 'N/A')}")
            
            return results
            
    except Exception as e:
        print(f"❌ ERRORE generale nell'esempio: {str(e)}")
        return []


def main():
    """Funzione principale con esempi corretti"""
    
    print("🔧 SAP SESSION MANAGER - VERSIONE CORRETTA")
    print("=" * 60)
    
    # Dizionario vuoto da popolare
    tempi_esecuzione = {}

    for n_thread in range(2, 7):
        try:
            start_time = time.time()
            # Test del Session Manager corretto
            esempio_consultazione_ordini(n_thread)
            
            end_time = time.time()
            tempo_impiegato = end_time - start_time

            # Salva il tempo nel dizionario
            tempi_esecuzione[n_thread] = tempo_impiegato
            
        except KeyboardInterrupt:
            print("\nOperazione interrotta dall'utente")
        except Exception as e:
            print(f"❌ ERRORE generale: {str(e)}")

    # Alla fine puoi visualizzare tutti i risultati
    print("\nRiepilogo tempi:")
    for threads, tempo in tempi_esecuzione.items():
        print(f"{threads} thread(s): {tempo:.2f} secondi")

if __name__ == "__main__":
    main()