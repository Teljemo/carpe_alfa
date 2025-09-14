from tasks import task_utils
from utils.excel_helpers import ExcelHelper
from threading import Lock

class TaskManager:
    def __init__(self):
        self.tasks = {}  # namn -> Task-objekt
        self.lock = Lock()
        self.excel = ExcelHelper()

    def add_task(self, name):
        with self.lock:
            if name not in self.tasks:
                self.tasks[name] = task_utils.Task(name)
            return self.tasks[name]

    def start_task(self, name):
        task = self.add_task(name)
        task.start()

    def pause_task(self, name):
        if name in self.tasks:
            self.tasks[name].pause()

    def stop_task(self, name):
        if name in self.tasks:
            task = self.tasks[name]
            task.stop()
            # Skriv till Excel-logg
            self.log_task(task)

    def get_elapsed(self, name):
        if name in self.tasks:
            return self.tasks[name].get_elapsed_str()
        return "00:00:00"

    # ------------------- Logga till Excel -------------------
    def log_task(self, task):
        """Lägger till en rad i lokal Excel med task info"""
        data = {
            "Task": task.name,
            "Elapsed": task.get_elapsed_str(),
            "StoppedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.excel.add_task_log(data)

    # ------------------- Hjälpfunktioner -------------------
    def list_tasks(self):
        return list(self.tasks.keys())
