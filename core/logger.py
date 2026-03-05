# File: core/logger.py

import queue
from datetime import datetime
from typing import Callable, Optional
from PyQt5.QtGui import QColor

class ThreadSafeLogger:
    """Sistema di logging thread-safe per operazioni parallele"""
    
    def __init__(self):
        self.log_queue = queue.Queue()
        self.gui_callback: Optional[Callable] = None
    
    def set_gui_callback(self, callback: Callable):
        """Imposta la callback per inviare log alla GUI"""
        self.gui_callback = callback
    
    def log(self, message: str, msg_type: str = "info"):
        """Aggiunge un messaggio alla coda (thread-safe)"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put((timestamp, message, msg_type))
    
    def info(self, message: str):
        self.log(message, "info")
    
    def success(self, message: str):
        self.log(message, "success")
    
    def error(self, message: str):
        self.log(message, "error")
    
    def warning(self, message: str):
        self.log(message, "warning")
    
    def debug(self, message: str):
        self.log(message, "debug")
    
    def process_queue(self):
        """Processa tutti i messaggi in coda (chiamato dal thread GUI)"""
        messages = []
        try:
            while True:
                timestamp, message, msg_type = self.log_queue.get_nowait()
                messages.append((timestamp, message, msg_type))
        except queue.Empty:
            pass
        return messages
    
    @staticmethod
    def get_color_for_type(msg_type: str) -> QColor:
        """Restituisce il colore appropriato per il tipo di messaggio"""
        colors = {
            "info": QColor("#0066CC"),
            "success": QColor("#008800"),
            "error": QColor("#CC0000"),
            "warning": QColor("#FF8800"),
            "debug": QColor("#888888")
        }
        return colors.get(msg_type, QColor("#000000"))