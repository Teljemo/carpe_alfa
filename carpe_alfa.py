# =================================================================================
#   Carpe Tempus – TIDRAPPORTERINGSAPPLIKATION
#   SKAPAD AV: EMANUEL TELJEMO
#   VERSION: 1.01
#
#
# --- Versionshistorik / Ändringslogg ---
#   V 1.01 - 2025-09-06
#   - Importer:
#       - import webbrowser
#   - Funktioner:
#       - Reparerat länk till hemsida i nederkant av appen, tidigare var den borttagen.
#       - Lagt till "encoding='utf-8'" vid alla filoperationer med JSON för att undvika problem med specialtecken.
# 
# =================================================================================
#   BESKRIVNING: En Tkinter-baserad applikation för att spåra ställtid och produktionstid
#   för olika artiklar. Appen loggar tider, antal och avvikelser till en Excel-fil lokalt
#   och synkroniserar data till en delad nätverksplats.
#   Flera användare kan köra appen samtidigt och dela artiklar/operationer filerna samt
#   filerna som synkroniseras till den delade mappen får unika namn med användarnamn och
#   tidsstämplar för att undvika konflikter.
#
#   =========================
#    HUVUDSEKTIONER:
#   - Importer
#   - Konstanter och sökvägar
#   - Konfigurationshantering
#   - Globala Variabler
#   - Hjälpfunktioner (tidsformatering m.m.)
#   - Backup och Synkronisering
#   - Filhantering (excel, artiklar, operationer)
#   - Laddning/Sparande av pågående jobb
#   - Systemstatusfunktioner
#   - UI-funktioner (skapande, uppdatering av artiklar/operationer, tidshantering)
#   - Avvikelsehantering
#   - Uppdateringsfunktioner för programmet
#   - Huvud-UI och Main loop
#   =========================
#
#   VIKTIGT: Ändra ALDRIG på någon kodrad, endast lägg till (svenska) kommentarer!
#            Lägg till linjer och sektioner tydligt. Påtala eventuella brister.
#            Om du ändrar sökvägar, filnamn, kolumner o.dyl – SE UPP så att beroenden
#            till dessa även uppdateras på andra ställen i filen!
# =================================================================================

# ---------------------------------------------------------------------------------
# [IMPORTER]
# Här importeras alla bibliotek som används i hela applikationen.
# Om du behöver utöka funktionaliteten eller vill ta bort något beroende,
# måste du kontrollera alla funktioner och UI-delar där modulens funktioner används!
# ---------------------------------------------------------------------------------
import tkinter as tk                    # För grafiskt gränssnitt (GUI)
from tkinter import ttk, messagebox     # Extra grafiska element och popup-rutor
import time                            # Tidshantering (t.ex. timestamps)
import datetime                        # Datum- och tidshantering
import pandas as pd                    # För läs/skriv av Excel-filer
import os                              # Fil- och sökvägshantering
import sys                             # Systemfunktioner (exempelvis script-path)
import psutil                          # Kontroll om processer körs (t.ex. Navision)
import pygetwindow as gw               # För att detektera/fokusera fönster (app/andra program)
import json                            # Läs och skriv av konfigurations- & datafiler i JSON-format
from pathlib import Path               # Modern filvägshantering
import shutil                          # Kopiera/flytta filer/kataloger
import logging                         # Loggning till fil och debugging-information
import webbrowser                      # Öppna webbläsare för uppdateringar

# ---------------------------------------------------------------------------------
# [KONSTANTER OCH SÖKVÄGAR]
# Här sätts kataloger, sökvägar och andra konstanter. Om man ändrar på en av dessa,
# måste man ofta också ändra relaterade delar längre ner! T.ex. kolumner, config-värden,
# filnamn, nätverkssökvägar.
# ---------------------------------------------------------------------------------

# Hämtar scriptets katalog så att alla interna filer alltid hamnar tillsammans
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Sätter upp loggning till fil. Alla loggar/varningar skrivs till time_tracker_log.txt.
LOG_FILE = os.path.join(SCRIPT_DIR, "time_tracker_log.txt")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ---------------------------------------------------------------------------------
# Funktion: show_splash_screen
# Visar en enkel startskärm med progressbar medan appen laddar nödvändiga filer.
# OBS! Koden kallar funktionen innan Tk-fönstret visas. Praktiskt för laddningstider.
# ---------------------------------------------------------------------------------
def show_splash_screen(root):
    """Visar en startskärm medan applikationen laddar i bakgrunden."""
    splash = tk.Toplevel(root)
    splash.title("Laddar...")
    splash.overrideredirect(True)  # Tar bort fönsterramen (ingen X-knapp)
    # Centrerar fönstret på skärmen
    splash_width = 300
    splash_height = 150
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (splash_width // 2)
    y = (screen_height // 2) - (splash_height // 2)
    splash.geometry(f'{splash_width}x{splash_height}+{x}+{y}')
    # Innehåll på startskärmen
    tk.Label(splash, text="Carpe Tempus", font=("Arial", 16, "bold"), pady=10).pack()
    status_label = tk.Label(splash, text="Läser in konfiguration...", font=("Arial", 10))
    status_label.pack()
    progress_bar = ttk.Progressbar(splash, orient="horizontal", length=200, mode="determinate")
    progress_bar.pack(pady=10)
    splash.update()  # Krävs för att progress bar etc. ska visas direkt
    return splash, status_label, progress_bar

# ---------------------------------------------------------------------------------
# [KONFIGURATIONSFILER]
# Den centrala konfigurationshanteringen.
# OBS! Om du lägger till en ny parameter här, måste du även uppdatera
# existerande config.json-filer – annars får användare defaultvärden!
# Om du lägger till nya sökvägar eller värden, se till att de används/uppdateras 
# korrekt i alla funktioner de berör!
# ---------------------------------------------------------------------------------
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

def load_config():
    """Laddar konfiguration från config.json. Skapar filen med standardvärden om den inte finns."""
    default_config = {
        # Sökväg till nätverkskatalog på SharePoint för sync
        "sharepoint_sync_path": os.path.join(r"C:\Users", os.getlogin(), r"Mercado Medic\Mercado Medic - MPAB\Shared Resources\data_sync\time_tracking"),
        "excel_file": "time_tracking_data.xlsx",           # Arbetsfil för loggning
        "running_tasks_file": "running_tasks.json",        # Håller all pågående jobb
        "last_backup_file": "last_backup.txt",             # Sista backupdatum
        "articles_file": "articles.xlsx",                  # Artikellista
        "articles_lock_file": ".articles.lock",            # Låsfil vid skrivning
        "operations_file": "operations.xlsx",              # Lista med operationer
        "operations_lock_file": ".operations.lock",        # Låsfil för operationer
        "temp_copy_path": os.path.join(SCRIPT_DIR, "temp/time_tracking_copies"),
        "daily_backup_path": os.path.join(os.getenv("APPDATA"), "BorjaStoppa"),
        "update_version_file": "version.json",
        # "deviation_codes": ["Väntetid", "Maskinfel", "Materialbrist", "Övrigt"]  # Avvikelser används nedan i avvikelsehantering!
    }

    # Om config-filen inte finns än (första gång) – skapa med defaultvärden
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
        return default_config

    else:
        # Läser existerande fil och slår ihop med default för att kompletta eventuella nya nycklar
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # Säkerställer att alla nycklar från default finns (för framtida utökning!)
            for key, value in default_config.items():
                if key not in config:
                    config[key] = value
            # Sparar tillbaka den uppdaterade configen (om nycklar lades till)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            return config
        except Exception as e:
            logging.error(f"Failed to load config file: {e}")
            # Brist: Om configen är trasig faller vi tillbaka på default utan att larma användaren
            #        Detta KAN innebära att data sparas på fel plats eller att nätverkssökvägar inte fungerar
            return default_config

# ---------------------------------------------------------------------------------
# [GLOBALA VARIABLER]
# Variabler som används genom hela programmet för att hålla applikationens tillstånd,
# lagrad data och annan viktig info.
# Var försiktig om du modifierar deras dataformat eller namn då många funktioner
# är beroende av just dessa namn och typer.
# ---------------------------------------------------------------------------------
data = []        # Huvudlistan där all loggad tid och status sparas (rad för rad)
articles = []    # Lista över alla artikelnummer (strängar)
operations = []  # Lista över operationer (strängar)
running_tasks = {}  # Dictionary över pågående arbetsuppgifter med artikelnummer som nyckel
selected_running_task_article = None  # Håller koll på vilken pågående uppgift som är vald i UI
app_start_time = time.time()          # Appens starttidpunkt (timestamp)
app_session_time = 0                  # Total tid appen varit det aktiva fönstret (i sekunder)
navision_session_time = 0             # Total tid Navision-processen körts (i sekunder)
last_app_check = time.time()          # Tidpunkt för senaste koll om appen var aktiv
last_navision_check = time.time()     # Tidpunkt för senaste koll om Navision var aktivt fönster
status_green_time = None              # Tidpunkt när SharePoint-nätverket blev "OK" första gången
access_green_time = None              # Tidpunkt när nätverksenheter blev "OK" första gången
computer_name = os.getenv("COMPUTERNAME")  # Datorns namn (från miljövariabel)
VERSION = "1.0"                      # Applikationens version. Viktigt för uppdateringar.

# OBSERVERA: Många funktioner räknar på dessa tider, om du ändrar namn måste du ändra
# i dessa funktioner för att undvika "NameError" eller felaktiga tidsberäkningar!

# ---------------------------------------------------------------------------------
# [HJÄLPFUNKTIONER FÖR TIDFORMAT OCH SÖKVÄG]
# Funktioner som hjälper till att konvertera sekunder till läsbara format eller decimaler.
# ---------------------------------------------------------------------------------
def format_time(seconds):
    """Formaterar sekunder till ett mer läsbart format som t.ex. '1h 30m 15s'.
    Om tiden är negativ returneras '0s'.
    Används för att visa tid i UI och loggar på ett tydligt sätt."""
    if seconds < 0:
        return "0s"
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    parts = []
    if hours > 0:
        parts.append(f"{hours}h")
    if minutes > 0 or (hours == 0 and seconds > 0):
        parts.append(f"{minutes}m")
    if seconds > 0:
        parts.append(f"{seconds}s")
    if not parts:
        return "0s"
    return " ".join(parts)

def get_decimal_time(seconds):
    """Omvandlar sekunder till minuter med två decimaler och rundar av.
    Används för att logga tider i minuter istället för sekunder."""
    return round(seconds / 60, 2)

def resource_path(relative_path):
    """Returnerar absolut sökväg till fil. Fungerar både när script körs som vanlig
    Python-fil och när det paketerats via PyInstaller (där _MEIPASS används)."""
    try:
        # PyInstaller temporär katalog
        base_path = sys._MEIPASS
    except Exception:
        # Vanlig körning normalt
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ---------------------------------------------------------------------------------
# [BACKUP OCH SYNKRONISERING]
# Funktioner för att skapa daglig backup av Excel-filen och för att kontrollera
# tillgång till nätverksresurser, t.ex. SharePoint och nätverksenheter.
# ---------------------------------------------------------------------------------
def perform_daily_backup():
    """Skapar en säker daglig backup av Excel-filen i en separat katalog.
    Om backup redan skapats idag görs inget. Viktigt för dataskydd!"""
    backup_path = config["daily_backup_path"]
    try:
        os.makedirs(backup_path, exist_ok=True)
        today = datetime.date.today().strftime("%Y-%m-%d")
        last_backup_date = None
        if os.path.exists(LAST_BACKUP_FILE):
            with open(LAST_BACKUP_FILE, 'r') as f:
                last_backup_date = f.read().strip()
        if last_backup_date != today:
            source_file = EXCEL_FILE
            backup_file_name = f"time_tracking_data_daglig_backup_{today}.xlsx"
            destination_file = os.path.join(backup_path, backup_file_name)
            if os.path.exists(source_file):
                shutil.copy2(source_file, destination_file)
                with open(LAST_BACKUP_FILE, 'w') as f:
                    f.write(today)
                logging.info(f"Daily backup created: {destination_file}")
            else:
                logging.warning(f"Could not find source file for daily backup: {source_file}")
    except Exception as e:
        logging.error(f"Failed to perform daily backup: {e}")

def check_system_access():
    """Kontrollerar nätverkstillgång genom att försöka skriva till och ta bort fil
    i SharePoint-mappen samt kontrollera att vissa nätverksenheter finns tillgängliga.
    Sätter tidsstämplar när status blir OK första gången.
    Returnerar tuple: (sharepoint_ok, all_drives_ok) som bool."""
    global status_green_time, access_green_time
    sharepoint_ok = False
    try:
        # Testar att skriva och ta bort fil i SharePoint-katalog
        temp_file = os.path.join(SHAREPOINT_SYNC_PATH, "temp_test.txt")
        with open(temp_file, 'w') as f:
            f.write("test")
        os.remove(temp_file)
        sharepoint_ok = True
    except (PermissionError, FileNotFoundError, OSError):
        sharepoint_ok = False

    # Kontrollera nätverksenheter (G:, Q:, R:) som måste finnas monterade
    drives = ['G:', 'Q:', 'R:']
    all_drives_ok = all(os.path.exists(drive) for drive in drives)

    # Spara tider för första "grön" status
    if sharepoint_ok and status_green_time is None:
        status_green_time = time.time()
    if all_drives_ok and access_green_time is None:
        access_green_time = time.time()

    return sharepoint_ok, all_drives_ok

def update_access_indicator():
    """Uppdaterar färgindikatorerna i UI (röda/gröna lampor) baserat på nätverkstillgång.
    Körs periodiskt var 30:e sekund."""
    sharepoint_ok, drives_ok = check_system_access()
    access_indicator.config(fg="green" if sharepoint_ok else "red")
    drive_indicator.config(fg="green" if drives_ok else "red")
    root.after(30000, update_access_indicator)  # Kör igen efter 30 sekunder

# ---------------------------------------------------------------------------------
# [FILHANTERING – LÄS OCH SKRIVNING AV EXCEL-FILER]
# Funktioner för att kontrollera, läsa in och spara data i Excel-filer.
# Dessa filer är kärnan i applikationens datalagring av artiklar,
# operationer och tidsrapporter.
# Var noggrann med att om ändringar görs i kolumnernas struktur eller namn,
# måste andra delar av programmet uppdateras för att undvika fel vid läsning/sprning.
# ---------------------------------------------------------------------------------

def check_excel_file():
    """Kontrollerar att huvud-Excel-filen för tidsrapportering finns och har
    rätt kolumner. Skapar filen med standardkolumner om den saknas eller är i fel format.
    Läser in befintlig data till global variabel `data`."""
    global data
    # Kolumnnamn som förväntas i Excel-filen. Viktigt att dessa stämmer överens med alla funktioner som läser/sparar.
    task_columns = ["Date", "User", "Article", "Task Type", "Task Start", "Task End",
                    "Setup Time (s)", "Production Time (s)", "Setup Parts", "Produced Parts",
                    "Scrapped Parts", "Time per Part (s)", "Note", "Operation",
                    "Task App Time (s)", "Task Navision Time (s)", "Computer",
                    "Extra Operators", "Setup Time (min)", "Production Time (min)",
                    "Time per Part (min)", "Deviation Time (min)", "Deviation Code"]

    if not os.path.exists(EXCEL_FILE):
        # Om filen saknas – skapa tom fil med rätt kolumner
        df = pd.DataFrame(columns=task_columns)
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Task_Data', index=False)
        logging.info(f"Created new Excel file: {EXCEL_FILE}")
    else:
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name='Task_Data')
            df["Article"] = df["Article"].astype(str)  # Säkerställer att artikelnummer är text
            # Kolla om filen har fel eller gamla kolumner
            missing_columns = [col for col in task_columns if col not in df.columns]
            extra_columns = [col for col in df.columns if col not in task_columns]
            if missing_columns or extra_columns:
                logging.warning(f"Old file format detected in {EXCEL_FILE}. Missing columns: {missing_columns}, Extra columns: {extra_columns}")
                # Lägg till saknade kolumner med defaultvärden
                for col in missing_columns:
                    if col in ["Operation", "Note", "Deviation Code"]:
                        df[col] = ""  # Strängkolumner får tomma strängar
                    else:
                        df[col] = 0   # Numeriska kolumner får 0 som default
                # Ta bort extra kolumner för att hålla format konsekvent
                df = df[task_columns]
            # Sparar inläst data till global lista som programmet använder
            data = df.values.tolist()
            logging.info(f"Loaded data from {EXCEL_FILE} with {len(data)} rows")
        except Exception as e:
            logging.error(f"Failed to load Task_Data from {EXCEL_FILE}: {e}")
            data = []
            # Om inläsning misslyckas – skapa en ny tom fil (risk för dataförlust!)
            df = pd.DataFrame(columns=task_columns)
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, sheet_name='Task_Data', index=False)

def check_articles_file():
    """Läser in artiklar från sin Excel-fil. Om filen inte finns skapas den.
    Laddar artiklar till global lista `articles`."""
    global articles
    articles.clear()
    if not os.path.exists(ARTICLES_FILE):
        # Skapar tom fil för artiklar om fil saknas
        df = pd.DataFrame(columns=["Article"])
        try:
            os.makedirs(SHAREPOINT_SYNC_PATH, exist_ok=True)
            with pd.ExcelWriter(ARTICLES_FILE, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Articles', index=False)
            logging.info(f"Created articles file: {ARTICLES_FILE}")
        except Exception as e:
            logging.error(f"Failed to create articles file {ARTICLES_FILE}: {e}")
    else:
        try:
            df = pd.read_excel(ARTICLES_FILE, sheet_name='Articles')
            # Säkerställer att artiklar är strängar, unika och sorterade
            articles.extend(sorted(df["Article"].dropna().astype(str).unique().tolist()))
            logging.info(f"Loaded {len(articles)} articles from {ARTICLES_FILE}")
        except Exception as e:
            logging.error(f"Failed to load articles from {ARTICLES_FILE}: {e}")
            articles.clear()

def check_operations_file():
    """Läser in operationer från sin Excel-fil. Om filen saknas skapas den.
    Laddar operationer till global lista `operations`."""
    global operations, selected_operation
    operations.clear()
    if not os.path.exists(OPERATIONS_FILE):
        df = pd.DataFrame(columns=["Operation"])
        try:
            os.makedirs(SHAREPOINT_SYNC_PATH, exist_ok=True)
            with pd.ExcelWriter(OPERATIONS_FILE, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Operations', index=False)
            logging.info(f"Created operations file: {OPERATIONS_FILE}")
        except Exception as e:
            logging.error(f"Failed to create operations file {OPERATIONS_FILE}: {e}")
    else:
        try:
            df = pd.read_excel(OPERATIONS_FILE, sheet_name='Operations')
            operations.extend(sorted(df["Operation"].dropna().astype(str).unique().tolist()))
            logging.info(f"Loaded {len(operations)} operations from {OPERATIONS_FILE}")
        except Exception as e:
            logging.error(f"Failed to load operations from {OPERATIONS_FILE}: {e}")
            operations.clear()

def save_articles_file():
    """Sparar artiklar till Excel och använder låsfil för att undvika konflikter vid samtidiga skrivningar.
    Om filen är låst försöker funktionen igen efter en stund. Om det misslyckas sparas en lokal temporär kopia."""
    df = pd.DataFrame(articles, columns=["Article"])
    max_attempts = 3
    delay_between_attempts = 5  # sekunder

    for attempt in range(max_attempts):
        if os.path.exists(ARTICLES_LOCK_FILE):
            logging.warning(f"Articles file locked: {ARTICLES_LOCK_FILE}. Waiting...")
            time.sleep(delay_between_attempts)
            continue
        try:
            with open(ARTICLES_LOCK_FILE, 'w') as f:
                f.write("locked")  # Skapar låsfil
            try:
                with pd.ExcelWriter(ARTICLES_FILE, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Articles', index=False)
                logging.info(f"Saved articles to {ARTICLES_FILE}")
                os.remove(ARTICLES_LOCK_FILE)  # Tar bort lås
                return True
            except Exception as e:
                os.remove(ARTICLES_LOCK_FILE)  # Tar bort lås även vid fel
                raise e
        except Exception as e:
            logging.warning(f"Attempt {attempt + 1}/{max_attempts} to save {ARTICLES_FILE} failed: {e}")
            if attempt < max_attempts - 1:
                time.sleep(delay_between_attempts)
    # Fallback: Spara lokal temporär kopia om nätverk inte är tillgängligt
    try:
        os.makedirs(TEMP_COPY_PATH, exist_ok=True)
        temp_filepath = os.path.join(TEMP_COPY_PATH, "articles.xlsx")
        with pd.ExcelWriter(temp_filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Articles', index=False)
        logging.info(f"Fallback: Saved articles to {temp_filepath}")
    except Exception as e:
        logging.error(f"Failed to save articles to {temp_filepath}: {e}")
    return False

def save_operations_file():
    """Sparar operationer till Excel med låsfil lika som för artiklar."""
    df = pd.DataFrame(operations, columns=["Operation"])
    max_attempts = 3
    delay_between_attempts = 5

    for attempt in range(max_attempts):
        if os.path.exists(OPERATIONS_LOCK_FILE):
            logging.warning(f"Operations file locked: {OPERATIONS_LOCK_FILE}. Waiting...")
            time.sleep(delay_between_attempts)
            continue
        try:
            with open(OPERATIONS_LOCK_FILE, 'w') as f:
                f.write("locked")
            try:
                with pd.ExcelWriter(OPERATIONS_FILE, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Operations', index=False)
                logging.info(f"Saved operations to {OPERATIONS_FILE}")
                os.remove(OPERATIONS_LOCK_FILE)
                return True
            except Exception as e:
                os.remove(OPERATIONS_LOCK_FILE)
                raise e
        except Exception as e:
            logging.warning(f"Attempt {attempt + 1}/{max_attempts} to save {OPERATIONS_FILE} failed: {e}")
            if attempt < max_attempts - 1:
                time.sleep(delay_between_attempts)
    try:
        os.makedirs(TEMP_COPY_PATH, exist_ok=True)
        temp_filepath = os.path.join(TEMP_COPY_PATH, "operations.xlsx")
        with pd.ExcelWriter(temp_filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Operations', index=False)
        logging.info(f"Fallback: Saved operations to {temp_filepath}")
    except Exception as e:
        logging.error(f"Failed to save operations to {temp_filepath}: {e}")
    return False

def save_data():
    """Sparar den insamlade tids- och produktionsdatan till huvud-Excel.
    Skapar samtidigt en tidsstämplad kopia på SharePoint. Tar bort gamla kopior för att spara utrymme."""
    df = pd.DataFrame(data, columns=["Date", "User", "Article", "Task Type", "Task Start", "Task End",
                                "Setup Time (s)", "Production Time (s)", "Setup Parts", "Produced Parts",
                                "Scrapped Parts", "Time per Part (s)", "Note", "Operation",
                                "Task App Time (s)", "Task Navision Time (s)", "Computer",
                                "Extra Operators", "Setup Time (min)", "Production Time (min)",
                                "Time per Part (min)", "Deviation Time (min)", "Deviation Code"])
    try:
        # Spara till lokal fil, ersätt 'Task_Data'-bladet med det nya datat
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Task_Data', index=False)
        logging.info(f"Saved data to {EXCEL_FILE}")

        # Skapar tidsstämplad kopia på nätverksplats för backup
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        username = os.getlogin()
        shared_file = os.path.join(SHAREPOINT_SYNC_PATH, f"time_tracking_data_{username}_{timestamp}.xlsx")
        shutil.copy2(EXCEL_FILE, shared_file)
        logging.info(f"Copied data to {shared_file}")

        # Rensar gamla kopior så bara 5 senaste finns kvar (som sparar utrymme)
        shared_files = [f for f in os.listdir(SHAREPOINT_SYNC_PATH)
                        if f.startswith("time_tracking_data_") and f.endswith(".xlsx")]
        if len(shared_files) > 5:
            shared_files.sort(key=lambda x: os.path.getmtime(os.path.join(SHAREPOINT_SYNC_PATH, x)))
            for old_file in shared_files[:-5]:
                os.remove(os.path.join(SHAREPOINT_SYNC_PATH, old_file))
                logging.info(f"Removed old copy: {old_file}")

    except Exception as e:
        logging.error(f"Failed to save data to {EXCEL_FILE} or copy to {SHAREPOINT_SYNC_PATH}: {e}")

# ---------------------------------------------------------------------------------
# [LADDA OCH SPARA PÅGÅENDE UPPGIFTER (RUNNING TASKS) - JSON-HANTERING]
# För att hålla reda på vilka jobb som pågår även då programmet startar om,
# sparas dessa i en JSON-fil. Dessa funktioner hanterar inläsning och sparning
# av dessa pågående jobb till/från fil.
# ---------------------------------------------------------------------------------

def load_running_tasks():
    """Laddar pågående uppgifter från JSON-fil när appen startar.
    Om filen inte finns eller är korrupt, hanteras fel och backupfil
    används om möjligt. Om allting misslyckas, erbjuds användaren att
    starta med tom listning."""
    global running_tasks
    if os.path.exists(RUNNING_TASKS_FILE):
        try:
            with open(config.get("running_tasks_file"), "r", encoding='utf-8') as f:
                running_tasks = json.load(f)
            logging.info(f"Loaded {len(running_tasks)} running tasks from {RUNNING_TASKS_FILE}")

            # Återställer tidräknare för varje pågående uppgift till nuvarande tid,
            # för att korrekt kunna räkna tid från senaste app-check.
            current_time = time.time()
            for task in running_tasks.values():
                task["last_app_check"] = current_time
                task["last_nav_check"] = current_time

        except (json.JSONDecodeError, Exception) as e:
            logging.error(f"Failed to load running tasks from {RUNNING_TASKS_FILE}: {e}")
            # Brist: Felmeddelandet hanteras bara i logg, användare får ingen detaljerad information
            # Försök återställningsfil från backup om finns
            if os.path.exists(TEMP_COPY_PATH):
                backup_file = os.path.join(TEMP_COPY_PATH, "running_tasks_backup.json")
                if os.path.exists(backup_file):
                    try:
                        with open(backup_file, 'r') as f:
                            running_tasks = json.load(f)
                        logging.info(f"Recovered running tasks from {backup_file}")
                    except Exception as e:
                        logging.error(f"Failed to recover from backup {backup_file}: {e}")
            # Frågar användaren om denne vill starta om med nya uppgifter eller lämna listan som tom
            if not messagebox.askyesno("Fel", "Kunde inte ladda pågående uppgifter. Vill du starta om med nya uppgifter?"):
                running_tasks = {}
            else:
                running_tasks = {}
    else:
        # Om fil saknas från start, starta med tom lista
        running_tasks = {}

def save_running_tasks():
    """Sparar alla pågående uppgifter till JSON-fil. Gör först en backup
    av befintlig fil för att kunna återställa vid problem.
    Om sparning misslyckas skrivs fel till logg."""
    try:
        Path(TEMP_COPY_PATH).mkdir(parents=True, exist_ok=True)
        backup_file = os.path.join(TEMP_COPY_PATH, "running_tasks_backup.json")
        if os.path.exists(RUNNING_TASKS_FILE):
            shutil.copy2(RUNNING_TASKS_FILE, backup_file)
        with open(config.get("running_tasks_file"), "w", encoding='utf-8') as f:
            json.dump(running_tasks, f, indent=4, ensure_ascii=False)
        logging.info(f"Saved running tasks to {RUNNING_TASKS_FILE}")
    except Exception as e:
        logging.error(f"Failed to save running tasks to {RUNNING_TASKS_FILE}: {e}")

# ---------------------------------------------------------------------------------
# [SYSTEMSTATUS: ÖVERVAKNING AV AKTIVITET, FÖNSTER OCH PROCESSER]
# Dessa funktioner jämför aktuellt aktivt fönster och processer för att spåra
# sessionstider för applikationen och Navision.
# Om dessa ändras måste tidräkning och UI-uppdateringar eventuellt uppdateras.
# ---------------------------------------------------------------------------------

def is_app_active():
    """Kontrollerar om "Carpe Tempus"-fönstret är det aktiva fönstret.
    Detta avgör om tidräkning för app-session ska räknas."""
    try:
        active_window = gw.getActiveWindow()
        if active_window and "Carpe Tempus" in active_window.title:
            return True
        return False
    except Exception:
        # Brist: Fel fångas brett, returnerar False utan felmeddelande
        return False

def is_navision_running():
    """Kontrollerar om Navision-klienten körs genom att leta efter processnamn.
    Om process finns, returnerar True."""
    try:
        for proc in psutil.process_iter(['name']):
            if "microsoft.dynamics.nav.client" in proc.info['name'].lower():
                return True
        return False
    except Exception:
        # Brist: Fel fångas brett, ingen logg här
        return False

def is_navision_active():
    """Kontrollerar om Navision-fönstret är det aktiva fönstret.
    Används för att köra tidräkning av uppgifter kopplade till Navision."""
    try:
        active_window = gw.getActiveWindow()
        if active_window and "Microsoft Dynamics NAV" in active_window.title:
            return True
        return False
    except Exception:
        # Brist: Fel fångas brett utan logg
        return False

# ---------------------------------------------------------------------------------
# [UPPDATERING AV TIMER-VARIABLER VARJE SEKUND]
# Funktion som uppdaterar sessionstider och tid på pågående uppgifter.
# Kallas rekursivt via root.after varje sekund.
# ---------------------------------------------------------------------------------

def update_timers():
    """Uppdaterar appens och Navisions sessionstider samt pågående jobbs tidräkningar.
    Tiden ökas endast om fönstret är aktivt (för app och Navision separat)."""
    global app_session_time, last_app_check, navision_session_time, last_navision_check
    current_time = time.time()

    # Uppdatera app-sessionstid om appens fönster är aktivt
    if is_app_active():
        app_session_time += current_time - last_app_check
    last_app_check = current_time

    # Uppdaterar Navision-sessionstid så länge process körs
    if is_navision_running():
        navision_session_time += current_time - last_navision_check
    last_navision_check = current_time

    # Uppdaterar tider för varje pågående jobb i running_tasks
    for task in running_tasks.values():
        # Tid kopplat till appens fönster aktivitet
        if is_app_active():
            task["task_app_time"] += current_time - task["last_app_check"]
        task["last_app_check"] = current_time

        # Tid kopplat till Navision-fönsteraktivitet
        if is_navision_active():
            task["task_navision_time"] += current_time - task["last_nav_check"]
        task["last_nav_check"] = current_time

    # Uppdatera UI för pågående jobb efter tidsuppdatering
    update_running_tasks_display()

    # Rekursivt kalla denna funktion igen efter 1 sekund
    root.after(1000, update_timers)

# ---------------------------------------------------------------------------------
# [UI-FUNKTIONER: ARTIKLAR OCH JOBB - SKAPA, UPPDATERA, STARTA, STOPPA]
# Dessa funktioner hanterar användarinteraktionen i gränssnittet.
# Här kan man skapa nya artiklar, starta och stoppa ställtid eller produktion,
# visa detaljer och hantera felhantering vid input.
# ---------------------------------------------------------------------------------

def create_new_article():
    """Skapar en ny artikel om textfältet inte är tomt och artikeln inte redan finns.
    Sparar sedan listan och uppdaterar UI.
    Varning visas om artikel redan finns eller fält är tomt."""
    new_article = new_article_var.get().strip()
    if new_article and new_article not in articles:
        articles.append(new_article)
        save_articles_file()
        update_article_list()
        new_article_var.set("")
    else:
        messagebox.showwarning("Varning", "Artikeln finns redan eller är tom.")

def refresh_articles():
    """Laddar om artiklar från fil och uppdaterar visningen.
    Kan användas om t.ex. filer ändrats externt."""
    check_articles_file()
    update_article_list()

def get_integer_input(prompt, title, min_value=0):
    """Visar en dialogruta för att mata in ett heltal.
    Låter användaren ange t.ex. antal delar.
    Validerar input och visar fel om värdet är fel.
    Returnerar int eller None om användaren avbryter."""
    result = None
    dialog = tk.Toplevel(root)
    dialog.title(title)
    # Positionerar dialogen centrerat över huvudfönstret
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    root_width = root.winfo_width()
    root_height = root.winfo_height()
    dialog_width = 350
    dialog_height = 150
    dialog.geometry(f"{dialog_width}x{dialog_height}+{root_x + int(root_width/2) - int(dialog_width/2)}+{root_y + int(root_height/2) - int(dialog_height/2)}")
    dialog.transient(root)  # Kopplar dialog till huvudfönstret
    dialog.grab_set()       # Gör dialogen modal

    tk.Label(dialog, text=prompt, pady=10).pack()
    entry = tk.Entry(dialog)
    entry.pack(pady=5)
    dialog.wait_visibility()
    entry.focus_set()

    def on_ok():
        nonlocal result
        try:
            input_text = entry.get().strip()
            if not input_text:
                messagebox.showerror("Fel", "Vänligen ange ett antal. Ange 0 om inga delar producerades.")
                entry.focus_set()
                return
            value = int(input_text)
            if value < min_value:
                messagebox.showerror("Fel", f"Värdet måste vara minst {min_value}.")
                entry.focus_set()
                return
            result = value
            dialog.destroy()
        except ValueError:
            messagebox.showerror("Fel", "Ange ett giltigt heltal.")
            entry.focus_set()

    def on_cancel():
        dialog.destroy()

    dialog.bind('<Return>', lambda event: on_ok())  # Enter-tangenten aktiverar OK

    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)
    ok_button = tk.Button(button_frame, text="OK", command=on_ok)
    ok_button.pack(side="left", padx=10)
    cancel_button = tk.Button(button_frame, text="Avbryt", command=on_cancel)
    cancel_button.pack(side="right", padx=10)
    dialog.wait_window()
    return result

def get_string_input(prompt, title):
    """Visar en dialogruta där användaren anger text.
    Returnerar den inmatade texten eller None vid avbryt."""
    result = ""
    dialog = tk.Toplevel(root)
    dialog.title(title)
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    root_width = root.winfo_width()
    root_height = root.winfo_height()
    dialog_width = 350
    dialog_height = 150
    dialog.geometry(f"{dialog_width}x{dialog_height}+{root_x + int(root_width/2) - int(dialog_width/2)}+{root_y + int(root_height/2) - int(dialog_height/2)}")
    dialog.transient(root)
    dialog.grab_set()

    tk.Label(dialog, text=prompt, pady=10).pack()
    entry = tk.Entry(dialog)
    entry.pack(pady=5)
    dialog.wait_visibility()
    entry.focus_set()

    def on_ok():
        nonlocal result
        result = entry.get()
        dialog.destroy()

    def on_cancel():
        nonlocal result
        result = None
        dialog.destroy()

    dialog.bind('<Return>', lambda event: on_ok())
    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)
    ok_button = tk.Button(button_frame, text="OK", command=on_ok)
    ok_button.pack(side="left", padx=10)
    cancel_button = tk.Button(button_frame, text="Avbryt", command=on_cancel)
    cancel_button.pack(side="right", padx=10)
    dialog.wait_window()
    return result

def start_setup():
    """Startar en ny ställtid på vald artikel.
    Om en pågående uppgift redan finns på samma artikel, varnar användaren.
    Skapar en "running_task" med typ 'setup'."""
    selected = article_selection_tree.selection()
    if not selected:
        messagebox.showwarning("Varning", "Välj en artikel först.")
        return
    article = article_selection_tree.item(selected)["values"][0]
    operation = selected_operation.get()
    extra_ops = selected_operators.get()
    if article in running_tasks:
        messagebox.showwarning("Varning", f"Artikeln '{article}' har redan en pågående uppgift. Ny ställtid kan inte startas.")
        return
    # Lägg till pågående jobb i dictionaryn
    running_tasks[article] = {
        "type": "setup",
        "start_time": time.time(),
        "last_app_check": time.time(),
        "last_nav_check": time.time(),
        "task_app_time": 0,
        "task_navision_time": 0,
        "operation": operation,
        "extra_operators": extra_ops
    }
    save_running_tasks()
    update_running_tasks_display()
    on_running_task_click(article)  # Markera automatiskt i UI

def start_task():
    """Startar en ny produktionstid.
    Om en ställtid pågår för samma artikel avslutas den först och loggas.
    Om en produktion redan pågår varnar användaren.
    Skapar nytt 'production' task i running_tasks."""
    selected = article_selection_tree.selection()
    if not selected:
        messagebox.showwarning("Varning", "Välj en artikel först.")
        return
    article = article_selection_tree.item(selected)["values"][0]
    operation = selected_operation.get()
    extra_ops = selected_operators.get()
    current_task = running_tasks.get(article)
    if current_task and current_task["type"] == "setup":
        # Avsluta ställtiden först
        end_time = time.time()
        task_time = end_time - current_task["start_time"]
        task_app_time = current_task["task_app_time"]
        task_navision_time = current_task["task_navision_time"]
        extra_ops = current_task["extra_operators"]
        setup_parts = get_integer_input("Ange antal ställtidsparts:", "Inmatning: Ställtid")
        if setup_parts is None:
            return  # Avbrutet
        note = get_string_input("Ange en notering (valfritt):", "Inmatning: Notering")
        if note is None:
            note = ""
        # Logga ställtidsraden i data
        data.append([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            os.getlogin(),
            str(article),
            "Setup",
            datetime.datetime.fromtimestamp(current_task["start_time"]).strftime("%Y-%m-%d %H:%M:%S"),
            datetime.datetime.fromtimestamp(end_time).strftime("%Y-%m-%d %H:%M:%S"),
            round(task_time, 2),
            0,
            setup_parts,
            0,
            0,
            0,
            note,
            current_task["operation"],
            round(task_app_time, 2),
            round(task_navision_time, 2),
            computer_name,
            extra_ops,
            get_decimal_time(task_time),  # Setup Time (min)
            0,  # Production Time (min)
            0,  # Time per Part (min)
            0,  # Deviation Time (min)
            ""
        ])
        del running_tasks[article]
        save_data()
    elif article in running_tasks:
        messagebox.showwarning("Varning", f"Artikeln '{article}' har redan en pågående uppgift. Ny produktion kan inte startas.")
        return
    # Skapa ny produktionstid
    running_tasks[article] = {
        "type": "production",
        "start_time": time.time(),
        "last_app_check": time.time(),
        "last_nav_check": time.time(),
        "task_app_time": 0,
        "task_navision_time": 0,
        "operation": operation,
        "extra_operators": extra_ops
    }
    save_running_tasks()
    update_running_tasks_display()
    on_running_task_click(article)

def stop_task():
    """Avslutar valt pågående jobb (ställtid eller produktion).
    Frågar användaren efter antal delar och notering innan loggning.
    Sparar data till huvud-loggen och uppdaterar UI."""
    global selected_running_task_article
    if not selected_running_task_article or selected_running_task_article not in running_tasks:
        messagebox.showwarning("Varning", "Välj ett pågående jobb först.")
        return
    article = selected_running_task_article
    end_time = time.time()
    task = running_tasks.pop(article)  # Ta bort från pågående jobs lista
    selected_running_task_article = None
    task_type = task["type"]
    task_time = end_time - task["start_time"]
    task_app_time = task["task_app_time"]
    task_navision_time = task["task_navision_time"]
    operation = task["operation"]
    extra_ops = task["extra_operators"]
    setup_time_s = round(task_time, 2) if task_type == "setup" else 0
    production_time_s = round(task_time, 2) if task_type == "production" else 0
    setup_parts = 0
    produced_parts = 0
    scrapped_parts = 0
    time_per_part_s = 0
    note = ""

    # Frågar efter parts beroende på typ av jobb
    if task_type == "setup":
        setup_parts = get_integer_input("Ange antal ställtidsparts:", "Inmatning: Ställtid")
        if setup_parts is None:
            running_tasks[article] = task  # Lägger tillbaka jobbet om avbrutet
            update_running_tasks_display()
            return
        note = get_string_input("Ange en notering (valfritt):", "Inmatning: Notering")
        if note is None:
            note = ""
    elif task_type == "production":
        produced_parts = get_integer_input("Ange antal producerade delar:", "Inmatning: Produktion")
        if produced_parts is None:
            running_tasks[article] = task
            update_running_tasks_display()
            return
        scrapped_parts = get_integer_input("Ange antal skrotade delar:", "Inmatning: Skrot")
        if scrapped_parts is None:
            running_tasks[article] = task
            update_running_tasks_display()
            return
        note = get_string_input("Ange en notering (valfritt):", "Inmatning: Notering")
        if note is None:
            note = ""
        good_parts = max(0, produced_parts - scrapped_parts)
        time_per_part_s = round(production_time_s / good_parts, 2) if good_parts > 0 else 0

    # Lägger till det avslutade jobbet i data-listan (huvudlogg)
    data.append([
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        os.getlogin(),
        str(article),
        task_type.capitalize(),
        datetime.datetime.fromtimestamp(task["start_time"]).strftime("%Y-%m-%d %H:%M:%S"),
        datetime.datetime.fromtimestamp(end_time).strftime("%Y-%m-%d %H:%M:%S"),
        setup_time_s,
        production_time_s,
        setup_parts,
        produced_parts,
        scrapped_parts,
        round(time_per_part_s, 2),
        note,
        operation,
        round(task_app_time, 2),
        round(task_navision_time, 2),
        computer_name,
        extra_ops,
        get_decimal_time(setup_time_s),
        get_decimal_time(production_time_s),
        get_decimal_time(time_per_part_s),
        0,  # Deviation Time (min)
        ""
    ])
    save_data()
    update_article_list()
    update_running_tasks_display()
    save_running_tasks()

# ---------------------------------------------------------------------------------
# [FUNKTIONER FÖR ATT UPPDATERA UI-LISTOR OCH VAL I GRÄNSSNITTET]
# ---------------------------------------------------------------------------------
def update_article_list():
    """Uppdaterar listan med artiklar som visas i UI, med eventuell sökfiltrering."""
    article_selection_tree.delete(*article_selection_tree.get_children())
    search_term = search_var.get().lower()
    for article in articles:
        if search_term in article.lower():
            article_selection_tree.insert("", "end", values=(article,))
    # Välj första artikel som standard om lista inte tom
    if article_selection_tree.get_children():
        article_selection_tree.selection_set(article_selection_tree.get_children()[0])
    operation_combobox['values'] = operations
    selected_operation.set(operations[0] if operations else "")

def show_article_details(event):
    """Visar historik/detaljer för vald artikel i respektive lista."""
    global selected_running_task_article
    selected = article_selection_tree.selection()
    if not selected:
        details_tree.delete(*details_tree.get_children())
        return
    article = article_selection_tree.item(selected)["values"][0]
    if article != selected_running_task_article:
        selected_running_task_article = None
        update_running_tasks_display()
    details_tree.delete(*details_tree.get_children())
    for row in data:
        if str(row[2]) == str(article):
            task_type = row[3]
            quantity = 0
            if task_type == "Setup":
                quantity = row[8]
            elif task_type == "Production":
                quantity = row[9]
            details_tree.insert("", "end", values=(
                row[0], row[1], row[2], row[13], quantity,
                format_time(row[6]), format_time(row[7]),
                row[3], f"{row[10]}", get_decimal_time(row[11])
            ))

def update_running_tasks_display():
    """Uppdaterar listan i UI över pågående jobb.
    Markerar valt jobb med blå bakgrund."""
    global selected_running_task_article
    for widget in running_frame.winfo_children():
        widget.destroy()
    tk.Label(running_frame, text=f"Användare: {os.getlogin()}", font=("Arial", 12, "bold")).pack(anchor="w")
    tk.Label(running_frame, text="Pågående jobb:", font=("Arial", 12, "bold")).pack(anchor="w")
    current_time = time.time()
    for article, task in running_tasks.items():
        elapsed = current_time - task["start_time"]
        elapsed_formatted = format_time(elapsed)
        row_frame = tk.Frame(running_frame, bg="lightblue" if article == selected_running_task_article else "SystemButtonFace")
        row_frame.pack(fill="x", pady=2)
        task_label = tk.Label(row_frame, text=f"{article} ({task['type'].capitalize()}, {task['operation']}): {elapsed_formatted} ",
                              cursor="hand2", bg=row_frame["bg"])
        task_label.pack(side="left", anchor="w", padx=5)
        task_label.bind("<Button-1>", lambda event, a=article: on_running_task_click(a))
    tk.Label(running_frame, text=f"App-sessionstid: {format_time(app_session_time)}").pack(anchor="w")
    tk.Label(running_frame, text=f"Navision-sessionstid: {format_time(navision_session_time)}").pack(anchor="w")

def on_running_task_click(article):
    """Hanterar klick på en pågående jobb-rad i UI. Växlar markering och väljer motsvarande artikel."""
    global selected_running_task_article
    if selected_running_task_article == article:
        selected_running_task_article = None  # Avmarkera om klickar igen
    else:
        selected_running_task_article = article
    select_article_in_treeview(article)
    update_running_tasks_display()

def select_article_in_treeview(article_name):
    """Markera och scrolla till en specifik artikel i artikellistan."""
    for item in article_selection_tree.get_children():
        if article_selection_tree.item(item, "values")[0] == article_name:
            article_selection_tree.selection_set(item)
            article_selection_tree.focus(item)
            article_selection_tree.see(item)
            return

# ---------------------------------------------------------------------------------
# [AVVIKELSEHANTERING]
# Användaren kan logga avvikelser kopplade till pågående jobb.
# Dialogruta visar val av avvikelsekod, notering och tid.
# ---------------------------------------------------------------------------------

def start_deviation():
    """Öppnar dialogruta för att logga en avvikelse på valt pågående jobb.
    Sparar avvikelsen i data-listan när användaren bekräftar."""
    global selected_running_task_article
    if not selected_running_task_article or selected_running_task_article not in running_tasks:
        messagebox.showwarning("Varning", "Välj ett pågående jobb först.")
        return
    result = {"code": None, "note": None, "time": None}
    dialog = tk.Toplevel(root)
    dialog.title("Avvikelse")
    dialog.geometry("350x250")
    dialog.grab_set()
    dialog.transient(root)

    tk.Label(dialog, text="Välj avvikelsekod:", font=("Arial", 10, "bold")).pack(pady=(10, 0))
    dev_code_var = tk.StringVar()
    dev_code_combobox = ttk.Combobox(dialog, textvariable=dev_code_var, values=config["deviation_codes"], state="readonly")
    dev_code_combobox.pack(fill="x", padx=20)
    if config["deviation_codes"]:
        dev_code_combobox.set(config["deviation_codes"][0])

    tk.Label(dialog, text="Ange en notering:", font=("Arial", 10, "bold")).pack(pady=(10, 0))
    dev_note_entry = tk.Entry(dialog, width=40)
    dev_note_entry.pack(fill="x", padx=20)

    tk.Label(dialog, text="Ange tid (timmar, minuter, sekunder):", font=("Arial", 10, "bold")).pack(pady=(10, 0))
    time_frame = tk.Frame(dialog)
    time_frame.pack()
    hours_var = tk.StringVar(value="0")
    minutes_var = tk.StringVar(value="0")
    seconds_var = tk.StringVar(value="0")
    tk.Label(time_frame, text="Timmar:").pack(side="left")
    hour_entry = tk.Entry(time_frame, textvariable=hours_var, width=5)
    hour_entry.pack(side="left", padx=5)
    tk.Label(time_frame, text="Minuter:").pack(side="left")
    min_entry = tk.Entry(time_frame, textvariable=minutes_var, width=5)
    min_entry.pack(side="left", padx=5)
    tk.Label(time_frame, text="Sekunder:").pack(side="left")
    sec_entry = tk.Entry(time_frame, textvariable=seconds_var, width=5)
    sec_entry.pack(side="left", padx=5)

    def on_ok():
        try:
            hours = int(hours_var.get())
            minutes = int(minutes_var.get())
            seconds = int(seconds_var.get())
            if hours < 0 or minutes < 0 or seconds < 0:
                messagebox.showerror("Fel", "Värdena måste vara positiva.")
                return
            if minutes > 59 or seconds > 59:
                messagebox.showerror("Fel", "Minuter och sekunder får inte vara större än 59.")
                return
            deviation_time = (hours * 3600) + (minutes * 60) + seconds
            if deviation_time == 0:
                messagebox.showerror("Fel", "Tiden måste vara större än 0.")
                return
            result["code"] = dev_code_var.get()
            result["note"] = dev_note_entry.get()
            result["time"] = deviation_time
            dialog.destroy()
        except ValueError:
            messagebox.showerror("Fel", "Ange giltiga heltal för tid.")

    def on_cancel():
        dialog.destroy()

    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=20)
    tk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=10)
    tk.Button(button_frame, text="Avbryt", command=on_cancel).pack(side="right", padx=10)

    dialog.wait_window()

    if result["code"] and result["time"]:
        article = selected_running_task_article
        operation = running_tasks[article]["operation"]
        log_deviation(article, operation, result["code"], result["note"], result["time"])

def log_deviation(article, operation, code, note, time_in_seconds):
    """Lägger till en rad i data-listan för en avvikelse med detaljer."""
    data.append([
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        os.getlogin(),
        str(article),
        "Avvikelse",
        "",
        "",
        0, 0, 0, 0, 0, 0,
        note,
        operation,
        0, 0,
        computer_name,
        0,
        0,  # Setup Time (min)
        0,  # Production Time (min)
        0,  # Time per Part (min)
        get_decimal_time(time_in_seconds),  # Deviation Time (min)
        code  # Deviation Code
    ])
    save_data()
    messagebox.showinfo("Avvikelse loggad", "Avvikelsen har loggats i Excel-filen.")

def on_closing():
    """Körs när användaren stänger fönstret. Sparar all data och städar upp."""
    if messagebox.askyesno("Bekräfta", "Är du säker på att du vill avsluta? All data sparas."):
        # Tar bort låsfiler om de finns kvar
        if os.path.exists(ARTICLES_LOCK_FILE):
            try:
                os.remove(ARTICLES_LOCK_FILE)
                logging.info(f"Removed lock file: {ARTICLES_LOCK_FILE}")
            except Exception as e:
                logging.error(f"Failed to remove lock file {ARTICLES_LOCK_FILE}: {e}")
        if os.path.exists(OPERATIONS_LOCK_FILE):
            try:
                os.remove(OPERATIONS_LOCK_FILE)
                logging.info(f"Removed lock file: {OPERATIONS_LOCK_FILE}")
            except Exception as e:
                logging.error(f"Failed to remove lock file {OPERATIONS_LOCK_FILE}: {e}")

        # Sparar allt i rätt ordning
        save_system_times()
        save_data()          # Sparar `data` till Excel
        save_articles_file()
        save_operations_file()
        perform_daily_backup()
        root.destroy()

def save_system_times():
    """Sparar information om sessionstider och tid till systemåtkomst i dataloggen."""
    global data, app_start_time, status_green_time, access_green_time, computer_name
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user = os.getlogin()
    # Loggar tid det tog att få kontakt med SharePoint
    if status_green_time is not None:
        time_to_status = round(status_green_time - app_start_time, 2)
        data.append([
            current_time, user, "", "Kontroll",
            datetime.datetime.fromtimestamp(app_start_time).strftime("%Y-%m-%d %H:%M:%S"),
            datetime.datetime.fromtimestamp(status_green_time).strftime("%Y-%m-%d %H:%M:%S"),
            0, 0, 0, 0, 0, 0,
            "Tid till Status",
            "", 0, 0,
            computer_name,
            0,
            0, 0, 0, 0, ""
        ])
    # Loggar tid det tog att få kontakt med nätverksenheter
    if access_green_time is not None:
        time_to_access = round(access_green_time - app_start_time, 2)
        data.append([
            current_time, user, "", "Kontroll",
            datetime.datetime.fromtimestamp(app_start_time).strftime("%Y-%m-%d %H:%M:%S"),
            datetime.datetime.fromtimestamp(access_green_time).strftime("%Y-%m-%d %H:%M:%S"),
            0, 0, 0, 0, 0, 0,
            "Tid till Åtkomst",
            "", 0, 0,
            computer_name,
            0,
            0, 0, 0, 0, ""
        ])


# ---------------------------------------------------------------------------------
# [HUVUD-UI OCH PROGRAMSTART]
# Här skapas Tkinter-fönstret, alla widgets byggs upp, 
# och timer-/statusuppdateringar startas.
# ---------------------------------------------------------------------------------

root = tk.Tk()  # Skapar huvudfönstret (roten till hela UI-trädet)
root.withdraw()  # Dölj huvudfönstret tills all initiering är klar (start med splash)

# --- Ladda konfiguration tidigt för att få rätt sökvägar ---
config = load_config()

# Sätt variabler för viktiga filvägar baserade på konfigurationen.
# Viktigt: Om dessa ändras måste de ändras på alla ställen i programmet där de används!
EXCEL_FILE = os.path.join(SCRIPT_DIR, config["excel_file"])
RUNNING_TASKS_FILE = os.path.join(SCRIPT_DIR, config["running_tasks_file"])
LAST_BACKUP_FILE = os.path.join(SCRIPT_DIR, config["last_backup_file"])
SHAREPOINT_SYNC_PATH = config["sharepoint_sync_path"]
ARTICLES_FILE = os.path.join(SHAREPOINT_SYNC_PATH, config["articles_file"])
ARTICLES_LOCK_FILE = os.path.join(SHAREPOINT_SYNC_PATH, config["articles_lock_file"])
OPERATIONS_FILE = os.path.join(SHAREPOINT_SYNC_PATH, config["operations_file"])
OPERATIONS_LOCK_FILE = os.path.join(SHAREPOINT_SYNC_PATH, config["operations_lock_file"])
TEMP_COPY_PATH = os.path.join(SCRIPT_DIR, "temp/time_tracking_copies")
DAILY_BACKUP_PATH = os.path.join(SHAREPOINT_SYNC_PATH, "save")
UPDATE_VERSION_FILE = os.path.join(SHAREPOINT_SYNC_PATH, config["update_version_file"])

# --- Visa laddningsskärm ---
splash_window, status_label, progress_bar = show_splash_screen(root)

try:
    # Läsa in data med progressupdates, visa i splash
    status_label.config(text="Läser in Excel-filer...")
    progress_bar['value'] = 25
    splash_window.update()
    check_excel_file()

    status_label.config(text="Läser in artiklar...")
    progress_bar['value'] = 50
    splash_window.update()
    check_articles_file()

    status_label.config(text="Läser in operationer...")
    progress_bar['value'] = 75
    splash_window.update()
    check_operations_file()

    status_label.config(text="Läser in pågående uppgifter...")
    progress_bar['value'] = 90
    splash_window.update()
    load_running_tasks()

    progress_bar['value'] = 100
    splash_window.update()

    time.sleep(1)  # Paus för visningseffekt

finally:
    splash_window.destroy()  # Döljer laddningsfönstret

root.deiconify()  # Visa huvudfönstret efter laddning klar

# --- Bygg UI ---

root.title("Carpe Tempus")  # Fönstertitel
root.geometry("1150x700")  # Standardstorlek
icon_path = resource_path("saxklocka.ico")  # Sökväg till ikon till fönster
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)  # Sätt ikon om den finns

root.protocol("WM_DELETE_WINDOW", on_closing)  # Koppla fönsterstängning till egen funktion

selected_operation = tk.StringVar()
selected_operators = tk.IntVar(value=0)

# Huvudram i fönstret som rymmer allt UI
main_frame = tk.Frame(root)
main_frame.pack(pady=10, padx=10, fill="both", expand=True)

# --- Input Area ---
input_frame = tk.Frame(main_frame)
input_frame.pack(fill="x")

tk.Label(input_frame, text="Sök artikel:").pack(side="left")

search_var = tk.StringVar()
def search_articles(*args):
    """Uppdaterar artikellistan baserat på sökfältet."""
    update_article_list()

search_var.trace("w", search_articles)  # Kör sökning varje gång text ändras
tk.Entry(input_frame, textvariable=search_var).pack(side="left", padx=5, fill="x", expand=True)

tk.Label(input_frame, text="Ny artikel:").pack(side="left")

new_article_var = tk.StringVar()
tk.Entry(input_frame, textvariable=new_article_var).pack(side="left", padx=5)

tk.Button(input_frame, text="Skapa artikel", command=create_new_article).pack(side="left")
tk.Button(input_frame, text="Uppdatera artiklar", command=refresh_articles).pack(side="left", padx=5)

# --- Sektion med två paneler: Artikellista och kontroller ---
split_frame = tk.Frame(main_frame)
split_frame.pack(fill="both", expand=True)

# Vänster panel visar artiklar
article_frame = tk.Frame(split_frame, width=500)
article_frame.pack(side="left", fill="both", expand=False)

tk.Label(article_frame, text="Tillgängliga artiklar:", font=("Arial", 12, "bold")).pack(anchor="w")

article_selection_frame = tk.Frame(article_frame)
article_selection_frame.pack(fill="both", expand=True)

article_selection_tree = ttk.Treeview(article_selection_frame, columns=("Article",), show="headings", height=10)
article_selection_tree.heading("Article", text="Artikel")
article_selection_tree.column("Article", width=200)

article_selection_tree.pack(side="left", fill="y")

scrollbar = ttk.Scrollbar(article_selection_frame, orient="vertical", command=article_selection_tree.yview)
scrollbar.pack(side="right", fill="y")

article_selection_tree.configure(yscrollcommand=scrollbar.set)
article_selection_tree.bind("<<TreeviewSelect>>", show_article_details)  # Visa detaljer vid val

# Höger panel med kontroller och pågående jobb
right_frame = tk.Frame(split_frame, width=500)
right_frame.pack(side="right", fill="both", expand=True)

selection_layout_frame = tk.Frame(right_frame)
selection_layout_frame.pack(fill="x", pady=0)

operation_frame = tk.Frame(selection_layout_frame)
operation_frame.pack(side="left", fill="x", expand=True)

tk.Label(operation_frame, text="Välj operation:", font=("Arial", 12, "bold")).pack(anchor="w")

operation_combobox = ttk.Combobox(operation_frame, textvariable=selected_operation, state="readonly")
operation_combobox.pack(fill="x")

operators_frame = tk.Frame(selection_layout_frame)
operators_frame.pack(side="left", fill="x", expand=True, padx=(10, 0))

tk.Label(operators_frame, text="Extra operatörer:", font=("Arial", 12, "bold")).pack(anchor="w")

operators_spinbox = ttk.Spinbox(operators_frame, from_=0, to=10, textvariable=selected_operators, state="readonly")
operators_spinbox.pack(fill="x")

# Knappsektion
button_frame = tk.Frame(right_frame)
button_frame.pack(fill="x", pady=10)

setup_button = tk.Button(button_frame, text="Starta ställtid", command=start_setup,
                         bg="blue", fg="white", font=("Arial", 12, "bold"), padx=10, pady=5)
setup_button.pack(side="left", padx=5)

start_button = tk.Button(button_frame, text="Starta produktion", command=start_task,
                         bg="green", fg="white", font=("Arial", 12, "bold"), padx=10, pady=5)
start_button.pack(side="left", padx=5)

stop_button = tk.Button(button_frame, text="Avsluta produktion", command=stop_task,
                        bg="red", fg="black", disabledforeground="white",
                        font=("Arial", 12, "bold"), padx=10, pady=5)
stop_button.pack(side="left", padx=5)

deviation_button = tk.Button(button_frame, text="Avvikelse", command=start_deviation,
                             bg="orange", fg="black", font=("Arial", 12, "bold"), padx=10, pady=5)
deviation_button.pack(side="left", padx=5)

# Status-indikatorer för nätverksåtkomst
tk.Label(button_frame, text="Status:", font=("Arial", 12, "bold")).pack(side="left", padx=5)
access_indicator = tk.Label(button_frame, text="●", font=("Arial", 20), fg="red")
access_indicator.pack(side="left", padx=5)

tk.Label(button_frame, text="Åtkomst:", font=("Arial", 12, "bold")).pack(side="left", padx=5)
drive_indicator = tk.Label(button_frame, text="●", font=("Arial", 20), fg="red")
drive_indicator.pack(side="left", padx=5)

# Ram för pågående jobb lista
running_frame = tk.Frame(right_frame)
running_frame.pack(fill="x", pady=10)

# Detaljer med historik för artikeln längst ner
tk.Label(main_frame, text="Artikeldetaljer:", font=("Arial", 12, "bold")).pack(anchor="w")

details_tree = ttk.Treeview(main_frame,
                            columns=("Date", "User", "Article", "Operation", "Quantity",
                                     "Setup Time (s)", "Production Time (s)", "Task Type",
                                     "Scrapped", "Time per Part (min)"),
                            show="headings")

details_tree.heading("Date", text="Datum")
details_tree.heading("User", text="Användare")
details_tree.heading("Article", text="Artikel")
details_tree.heading("Operation", text="Operation")
details_tree.heading("Quantity", text="Antal")
details_tree.heading("Setup Time (s)", text="Ställtid")
details_tree.heading("Production Time (s)", text="Produktionstid")
details_tree.heading("Task Type", text="Typ av jobb")
details_tree.heading("Scrapped", text="Skrotade")
details_tree.heading("Time per Part (min)", text="Tid/del (min)")

# Justera kolumnbredder
details_tree.column("Date", width=110)
details_tree.column("User", width=150)
details_tree.column("Article", width=80)
details_tree.column("Operation", width=150)
details_tree.column("Quantity", width=50)
details_tree.column("Setup Time (s)", width=80)
details_tree.column("Production Time (s)", width=80)
details_tree.column("Task Type", width=100)
details_tree.column("Scrapped", width=80)
details_tree.column("Time per Part (min)", width=80)

details_tree.pack(fill="both", expand=True)

# --- Informationsfält och version längst ner ---
info_frame = tk.Frame(root)
info_frame.pack(side="bottom", fill="x", padx=10, pady=5)


# ---- Länk till skaparen
def open_link(url):
    # Kontrollerar om URL:en redan har ett protokoll (http/https)
    if not url.startswith(('http://', 'https://')):
        url = 'http://' + url  # Lägger till http:// om det saknas
    try:
        webbrowser.open(url, new=2)
    except Exception as e:
        logging.error(f"Kunde inte öppna länk {url}: {e}")
        messagebox.showerror("Länkfel", f"Kunde inte öppna länken: {url}")

# ---- Hämta länken från konfigurationen
# Vi använder .get() med en fallback till "www.google.se" ifall nyckeln saknas eller är tom
made_by_link_url = config.get("made_by_link", "www.google.se") # Fallback om länken saknas eller är tom
print("DEBUG: made_by_link_url =", made_by_link_url)

made_by_label = tk.Label(info_frame, text="© 2025, Emanuel Teljemo, Alla rättigheter förbehållna.", fg="blue", cursor="hand2")
made_by_label.pack(side="left", padx=10)

# ---- Bind klicket till funktionen, och använd den hämtade URL:en
# Vi kontrollerar att länken inte är tom eller bara vitutrymme innan vi binder
if made_by_link_url and made_by_link_url.strip():
    made_by_label.bind("<Button-1>", lambda e: open_link(made_by_link_url))
else:
    # Om länken saknas eller är tom, gör inget vid klick
    pass

# Visning av version längst ner intill länken
tk.Label(info_frame, text=f"Version: {VERSION}").pack(side="left")

# Uppdateringskontroll knapp
def check_for_updates():
    """Kontrollerar om en ny version av programmet finns tillgänglig och informerar användaren."""
    try:
        if os.path.exists(UPDATE_VERSION_FILE):
            with open(UPDATE_VERSION_FILE, "r") as f:
                version_info = json.load(f)
            latest_version = version_info.get("latest_version", VERSION)
            update_url = version_info.get("update_url", None)
            if latest_version != VERSION:
                msg = f"En ny version ({latest_version}) finns tillgänglig."
                if update_url:
                    msg += f"\nLadda ner från: {update_url}"
                messagebox.showinfo("Uppdatering tillgänglig", msg)
            else:
                messagebox.showinfo("Uppdatering", "Du har redan den senaste versionen.")
        else:
            messagebox.showwarning("Uppdatering", "Ingen versionsinformation hittades.")
    except Exception as e:
        messagebox.showerror("Fel", f"Kunde inte kontrollera uppdatering: {e}")

update_button = tk.Button(info_frame, text="Kontrollera Uppdatering", command=check_for_updates)
update_button.pack(side="right", padx=10)

# --- Starta appens funktioner och UI-uppdateringar ---

update_article_list()  # Visa lista med artiklar
operation_combobox['values'] = operations
selected_operation.set(operations[0] if operations else "")

update_running_tasks_display()  # Visa pågående jobb

# Starta periodisk kontroll och UI-uppdatering
root.after(1000, update_access_indicator)  # Starta kontroll av systemtillgång var 30:e sekund
update_timers()  # Starta huvudtimer som uppdaterar sekunder varje sekund

# Huvudloopen som håller appen igång
root.mainloop()
