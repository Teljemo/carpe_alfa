import os
import time
import threading
from datetime import datetime
from config.settings import config
from utils import file_helpers as fh

class StatusMonitor:
    def __init__(self):
        self.shared_folder = fh.safe_path(config.shared_folder)
        self.disk_list = config.disk_check_list
        self.navision_interval = config.navision_check_interval
        self.app_start_time = datetime.now()
        self.access_time = None
        self.status = False
        self.access = False
        self.app_active = False
        self.nav_active = False
        self._stop_thread = False
        self.lock = threading.Lock()

        # Starta bakgrundstråd
        self.thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.thread.start()

    # ------------------- Monitorloop -------------------
    def monitor_loop(self):
        while not self._stop_thread:
            self.check_status()
            self.check_access()
            self.check_app_activity()
            self.check_navision_activity()
            time.sleep(1)

    def stop(self):
        self._stop_thread = True
        self.thread.join()

    # ------------------- Statuslampor -------------------
    def check_status(self):
        """Kontrollerar om delad mapp finns"""
        self.status = os.path.exists(self.shared_folder)

    def check_access(self):
        """Kontrollerar om alla diskar finns"""
        self.access = all(os.path.exists(disk) for disk in self.disk_list)
        if self.access and self.access_time is None:
            self.access_time = datetime.now()

    # ------------------- Timers -------------------
    def elapsed_app_start(self):
        """Tid från appstart"""
        return (datetime.now() - self.app_start_time).total_seconds()

    def elapsed_access(self):
        """Tid från appstart tills åtkomst erhållen"""
        if self.access_time:
            return (datetime.now() - self.access_time).total_seconds()
        return None

    # ------------------- Aktivitet -------------------
    def check_app_activity(self):
        """Kontrollerar om appen är aktiv (fokuserat fönster)"""
        try:
            import win32gui, win32process
            fg_window = win32gui.GetForegroundWindow()
            thread_id, proc_id = win32process.GetWindowThreadProcessId(fg_window)
            self.app_active = (proc_id == os.getpid())
        except ImportError:
            self.app_active = True  # fallback

    def check_navision_activity(self):
        """Kontrollerar om Navision körs och är aktivt"""
        try:
            import psutil
            navision_processes = [p for p in psutil.process_iter(['name']) if 'NAV' in p.info['name'].upper()]
            self.nav_active = any(navision_processes)
        except ImportError:
            self.nav_active = False

    # ------------------- Hjälpfunktioner -------------------
    def get_status(self):
        return self.status

    def get_access(self):
        return self.access

    def get_elapsed_app_start(self):
        return self.elapsed_app_start()

    def get_elapsed_access(self):
        return self.elapsed_access()

    def is_app_active(self):
        return self.app_active

    def is_nav_active(self):
        return self.nav_active
