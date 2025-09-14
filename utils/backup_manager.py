from config import data_storage
from utils import file_helpers as fh
import threading
import time

class BackupManager:
    def __init__(self, interval_seconds=3600):
        self.storage = data_storage.DataStorage()
        self.interval = interval_seconds  # kör backup var X sekunder (för realtidskopiering)
        self._stop_thread = False
        self.thread = threading.Thread(target=self._backup_loop, daemon=True)
        self.thread.start()

    # ------------------- Loop -------------------
    def _backup_loop(self):
        while not self._stop_thread:
            self.copy_local_to_shared()
            time.sleep(self.interval)

    def stop(self):
        """Stoppar backup-tråden"""
        self._stop_thread = True
        self.thread.join()

    # ------------------- Backup-funktioner -------------------
    def copy_local_to_shared(self):
        """Skapar timestamp-kopia i delad mapp"""
        try:
            self.storage.copy_to_shared()
        except Exception as e:
            print(f"Backup to shared failed: {e}")

    def daily_backup_local(self):
        """Skapar daglig backup av lokal fil"""
        try:
            self.storage.daily_backup()
        except Exception as e:
            print(f"Daily local backup failed: {e}")

    def daily_backup_articles(self):
        """Skapar daglig backup av articles.xlsx"""
        try:
            self.storage.backup_articles()
        except Exception as e:
            print(f"Daily articles backup failed: {e}")
