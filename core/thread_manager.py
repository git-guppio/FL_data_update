# File: core/thread_manager.py

import concurrent.futures
import threading
from typing import List, Callable, Any, Optional, Dict
from core.base_component import BaseComponent

class ThreadManager(BaseComponent):
    """Gestisce l'esecuzione parallela di operazioni"""
    
    def __init__(self, session_manager, logger=None):
        self.session_manager = session_manager
        super().__init__(logger)
    
    def execute_parallel_operations(
        self,
        operation_func: Callable,
        data_list: List[Any],
        SAP_data_directory: str,
        max_workers: Optional[int] = None
        ) -> List[Dict]:
        """
        Esegue operazioni in parallelo
        
        Args:
            operation_func: Funzione da eseguire (riceve session e data)
            data_list: Lista di dati da processare
            SAP_data_directory: Directory per salvare i dati SAP
            max_workers: Numero di worker (default: min(len(data_list), max_sessions))
            
        Returns:
            Lista di risultati (dizionari)
        
        Example:
            >>> results = thread_manager.execute_parallel_operations(
            ...     operation_func=self.sap_operations.consulta_ordine,
            ...     data_list=['ORD001', 'ORD002', 'ORD003'],
            ...     SAP_data_directory="./data",
            ...     max_workers=3
            ... )            
        """
        if max_workers is None:
            max_workers = min(len(data_list), self.session_manager.max_sessions)
        
        # ✅ CORRETTO: Usa self.log() invece di if self.logger:
        self.log(f"Avvio {len(data_list)} operazioni con {max_workers} worker", "info")
        
        def execute_with_session(data, SAP_data_directory):
            """Wrapper thread-safe"""
            with self.session_manager.get_session() as session:
                if session:
                    try:
                        return operation_func(session, data, SAP_data_directory)
                    except Exception as e:
                        # ✅ CORRETTO
                        self.log(f"Errore operazione: {str(e)}", "error")
                        return {
                            'impianto': str(data),
                            'data': None,
                            'status': 'error',
                            'errore': str(e)
                        }
                else:
                    # ✅ CORRETTO
                    self.log("Sessione non disponibile", "error")
                    return {
                        'impianto': str(data),
                        'data': None,
                        'status': 'error',
                        'errore': 'Sessione non disponibile'
                    }
        
        results = []
        # Esegui in parallelo
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=max_workers,
            thread_name_prefix="SAP_Worker"
        ) as executor:
            # Invia tutte le operazioni
            futures = [
                executor.submit(execute_with_session, data, SAP_data_directory) 
                for data in data_list
            ]
            
            # Raccogli i risultati
            for i, future in enumerate(concurrent.futures.as_completed(futures)):
                try:
                    result = future.result(timeout=60)
                    if result:
                        results.append(result)
                        status = "✅" if result.get('status') == 'success' else "❌"
                        # ✅ CORRETTO
                        self.log(
                            f"{status} Operazione {i+1}/{len(data_list)} - "
                            f"data: {result.get('impianto')}", 
                            "info"
                        )
                        
                except concurrent.futures.TimeoutError:
                    # ✅ CORRETTO
                    self.log(f"Timeout operazione {i+1}/{len(data_list)}", "warning")
                    
                except Exception as e:
                    # ✅ CORRETTO
                    self.log(f"Errore operazione {i+1}: {str(e)}", "error")
        
        # ✅ CORRETTO
        self.log(f"Completate {len(results)}/{len(data_list)} operazioni", "success")
        
        return results