import tkinter as tk
from tkinter import ttk
from threading import Thread
import time

class StatusLamp(tk.Label):
    """LED-lampa som ändrar färg beroende på status"""
    def __init__(self, master, text="", **kwargs):
        super().__init__(master, text=text, width=2, **kwargs)
        self.status = False
        self.update_color()

    def set_status(self, status: bool):
        self.status = status
        self.update_color()

    def update_color(self):
        color = "green" if self.status else "red"
        self.config(bg=color)

class TimerLabel(tk.Label):
    """Visar en timer i HH:MM:SS-format"""
    def __init__(self, master, get_seconds_func, **kwargs):
        super().__init__(master, text="00:00:00", **kwargs)
        self.get_seconds = get_seconds_func
        self._running = True
        self.thread = Thread(target=self.update_loop, daemon=True)
        self.thread.start()

    def update_loop(self):
        while self._running:
            seconds = self.get_seconds()
            if seconds is not None:
                h, rem = divmod(int(seconds), 3600)
                m, s = divmod(rem, 60)
                self.config(text=f"{h:02d}:{m:02d}:{s:02d}")
            time.sleep(1)

    def stop(self):
        self._running = False

class TaskRow(tk.Frame):
    """En rad i task-listan: namn + start/pause/stop-knappar"""
    def __init__(self, master, task_name, manager, **kwargs):
        super().__init__(master, **kwargs)
        self.task_name = task_name
        self.manager = manager

        self.name_label = tk.Label(self, text=task_name, width=20, anchor="w")
        self.start_btn = ttk.Button(self, text="Start", width=6, command=self.start_task)
        self.pause_btn = ttk.Button(self, text="Pause", width=6, command=self.pause_task)
        self.stop_btn = ttk.Button(self, text="Stop", width=6, command=self.stop_task)

        self.name_label.pack(side=tk.LEFT, padx=(2,5))
        self.start_btn.pack(side=tk.LEFT, padx=2)
        self.pause_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn.pack(side=tk.LEFT, padx=2)

    def start_task(self):
        self.manager.start_task(self.task_name)

    def pause_task(self):
        self.manager.pause_task(self.task_name)

    def stop_task(self):
        self.manager.stop_task(self.task_name)
