import tkinter as tk
from tkinter import ttk
from gui.widgets import StatusLamp, TimerLabel, TaskButtonFrame
from utils.timers import StatusMonitor
from tasks.task_manager import TaskManager

class MainWindow:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.task_manager = TaskManager()
        self.status_monitor = StatusMonitor()

        # ------------------- Statuslampor -------------------
        self.status_frame = tk.Frame(root)
        self.status_frame.pack(pady=5)
        self.status_label = tk.Label(self.status_frame, text="Status:")
        self.status_label.pack(side=tk.LEFT, padx=5)
        self.status_lamp = StatusLamp(self.status_frame)
        self.status_lamp.pack(side=tk.LEFT, padx=5)
        self.access_label = tk.Label(self.status_frame, text="Åtkomst:")
        self.access_label.pack(side=tk.LEFT, padx=5)
        self.access_lamp = StatusLamp(self.status_frame)
        self.access_lamp.pack(side=tk.LEFT, padx=5)

        # ------------------- Timers -------------------
        self.timer_frame = tk.Frame(root)
        self.timer_frame.pack(pady=5)
        self.app_timer_label = TimerLabel(self.timer_frame, get_seconds_func=self.status_monitor.get_elapsed_app_start)
        self.app_timer_label.pack(side=tk.LEFT, padx=5)
        self.access_timer_label = TimerLabel(self.timer_frame, get_seconds_func=self.status_monitor.get_elapsed_access)
        self.access_timer_label.pack(side=tk.LEFT, padx=5)

        # ------------------- Task lista -------------------
        self.task_frame = tk.Frame(root)
        self.task_frame.pack(pady=10, fill=tk.X)
        self.task_entries = {}
        self.refresh_tasks()

        # Starta UI-loop för statuslampor
        self.root.after(1000, self.update_status)

    def add_task_ui(self, task_name):
        frame = TaskButtonFrame(self.task_frame, task_name, self.task_manager)
        frame.pack(fill=tk.X, pady=2)
        self.task_entries[task_name] = frame

    def refresh_tasks(self):
        for task_name in self.task_manager.list_tasks():
            if task_name not in self.task_entries:
                self.add_task_ui(task_name)

    def update_status(self):
        self.status_lamp.set_status(self.status_monitor.get_status())
        self.access_lamp.set_status(self.status_monitor.get_access())
        self.root.after(1000, self.update_status)
