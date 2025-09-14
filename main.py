# main.py
import tkinter as tk
from app.ui import build_ui
from app.utils import load_config


def main():
    # Ladda config
    config = load_config("config.json")

    # Skapa huvudfönstret
    root = tk.Tk()
    root.title("Carpe Tempus")
    root.geometry(config.get("window_geometry", "1150x700"))

    # Bygg hela UI:t (layout + komponenter)
    build_ui(root, config)

    # Starta appen
    root.mainloop()


if __name__ == "__main__":
    main()
