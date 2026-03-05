# File: core/base_component.py

class BaseComponent:
    """
    Classe base per tutti i componenti con logging
    
    Fornisce un metodo unificato self.log() per logging thread-safe
    """
    
    def __init__(self, logger=None):
        """
        Args:
            logger: Istanza di ThreadSafeLogger (opzionale)
        """
        self.logger = logger
    
    def log(self, message: str, msg_type: str = "info"):
        """
        Helper per logging thread-safe
        
        Usa questo metodo per loggare da qualsiasi classe
        
        Args:
            message: Messaggio da loggare
            msg_type: Tipo di messaggio (info, success, error, warning, debug)
        
        Example:
            self.log("Operazione completata", "success")
            self.log("Attenzione", "warning")
            self.log("Errore critico", "error")
        """
        if self.logger:
            self.logger.log(message, msg_type)
        else:
            # ✅ Fallback su print se logger non disponibile
            print(f"[{msg_type.upper()}] {message}")            