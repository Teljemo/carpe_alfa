import tkinter as tk
from config.settings import config
from gui.main_window import MainWindow
from utils.backup_manager import BackupManager

def main():
    root = tk.Tk()
    root.title("Carpe Alfa")

    # Starta backup-manager
    backup_manager = BackupManager(interval_seconds=3600)  # realtidskopiering varje timme

    app = MainWindow(root, config)

    def on_close():
        # Stoppa bakgrundstrådar
        app.status_monitor.stop()
        app.app_timer_label.stop()
        app.access_timer_label.stop()
        backup_manager.stop()

        # Kör dagliga backup
        backup_manager.daily_backup_local()
        backup_manager.daily_backup_articles()

        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()
