import os
from datetime import datetime
import getpass

def current_user():
    """Hämtar aktuellt Windows-användarnamn"""
    return getpass.getuser()

def timestamp(fmt="%Y-%m-%d_%H-%M-%S"):
    """Returnerar aktuell tid som sträng"""
    return datetime.now().strftime(fmt)

def dated_string(fmt="%Y-%m-%d"):
    """Returnerar dagens datum som sträng"""
    return datetime.now().strftime(fmt)

def safe_path(path):
    """Returnerar absolut sökväg"""
    return os.path.abspath(os.path.expanduser(path))
