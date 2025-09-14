import tkinter as tk
from tkinter import ttk
from gui.widgets import StatusLamp, TimerLabel, TaskRow
from utils.timers import StatusMonitor
from tasks.task_manager import TaskManager

class MainWindow:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.task_manager = TaskManager()
        self.status_monitor = StatusMonitor()

        # -------- ÖVERSTA RADEN (status + timers) ----------
        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        tk.Label(top_frame, text="Status:").pack(side=tk.LEFT, padx=2)
        self.status_lamp = StatusLamp(top_frame)
        self.status_lamp.pack(side=tk.LEFT, padx=2)
        tk.Label(top_frame, text="Åtkomst:").pack(side=tk.LEFT, padx=2)
        self.access_lamp = StatusLamp(top_frame)
        self.access_lamp.pack(side=tk.LEFT, padx=2)

        self.app_timer_label = TimerLabel(top_frame, get_seconds_func=self.status_monitor.get_elapsed_app_start)
        self.app_timer_label.pack(side=tk.LEFT, padx=10)
        self.access_timer_label = TimerLabel(top_frame, get_seconds_func=self.status_monitor.get_elapsed_access)
        self.access_timer_label.pack(side=tk.LEFT, padx=10)

        # -------- HUVUDPANEL MED VÄNSTER + HÖGER ----------
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- Vänsterpanel: scrollad task-lista ---
        left_frame = tk.Frame(main_frame, width=300)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.canvas = tk.Canvas(left_frame, borderwidth=0)
        self.task_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.canvas.yview)
        self.task_container = tk.Frame(self.canvas)

        self.task_container.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0,0), window=self.task_container, anchor="nw")
        self.canvas.configure(yscrollcommand=self.task_scroll.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.Y, expand=True)
        self.task_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Scroll med mushjul fungerar även utanför canvas
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.task_rows = {}
        self.refresh_tasks()

        # --- Högerpanel: detaljer/visning ---
        self.right_frame = tk.Frame(main_frame, bg="#f0f0f0")
        self.right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Starta uppdatering av statuslampor
        self.root.after(500, self.update_status)

    # ---------- FUNKTIONER ----------
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def add_task_ui(self, task_name):
        row = TaskRow(self.task_container, task_name, self.task_manager)
        row.pack(fill=tk.X, pady=1)
        self.task_rows[task_name] = row

    def refresh_tasks(self):
        for name in self.task_manager.list_tasks():
            if name not in self.task_rows:
                self.add_task_ui(name)

    def update_status(self):
        self.status_lamp.set_status(self.status_monitor.get_status())
        self.access_lamp.set_status(self.status_monitor.get_access())
        self.root.after(1000, self.update_status)
