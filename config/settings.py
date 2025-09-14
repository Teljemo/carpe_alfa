import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'config.json')

class Config:
    def __init__(self, path=CONFIG_PATH):
        self.path = path
        self.load()

    def load(self):
        with open(self.path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self.local_data_path = data.get("local_data_path", "time_tracking_data.xlsx")
        self.shared_folder = data.get("shared_folder", r"\\SERVER\Shared")
        self.backup_folder = data.get("backup_folder", r"\\SERVER\Backup")
        self.articles_file = data.get("articles_file", "articles.xlsx")
        self.operations_file = data.get("operations_file", "operations.xlsx")
        self.navision_check_interval = data.get("navision_check_interval", 10)
        self.disk_check_list = data.get("disk_check_list", ["D:", "E:", "F:"])
        self.backup_time_format = data.get("backup_time_format", "%Y-%m-%d_%H-%M-%S")
        self.backup_date_format = data.get("backup_date_format", "%Y-%m-%d")

config = Config()
