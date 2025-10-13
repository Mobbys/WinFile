import sys
from cx_Freeze import setup, Executable
import matplotlib # Importato per trovare la sua cartella dati

# --- Dati dell'applicazione ---
APP_NAME = "WinFile"
APP_VERSION = "4.7"

# --- Opzioni di Base ---
# Queste sono le librerie che il tuo programma usa.
packages = [
    "customtkinter",
    "tkinterdnd2",
    "PIL",
    "fitz", # PyMuPDF
    "reportlab",
    "requests",
    "packaging",
    "tkinter", # Aggiungiamo tkinter esplicitamente per sicurezza
    "numpy",
    "cv2",
    "ezdxf",
    "matplotlib"
]

# --- NUOVO: Trova e include automaticamente la cartella dati di matplotlib ---
matplotlib_data_path = matplotlib.get_data_path()
include_files = [
    ('apps', 'apps'),
    ('config.json', 'config.json'),
    (matplotlib_data_path, 'mpl-data') # Copia la cartella dati nella root della build
]

# --- Configurazione per l'eseguibile ---
base = None
if sys.platform == "win32":
    base = "Win32GUI" # Nasconde la finestra del terminale all'avvio

# --- Costruisci il nome della cartella di output ---
# Questa riga crea una stringa tipo "build/WinFile 4.0"
build_dir_name = f"build/{APP_NAME} {APP_VERSION}"


# Opzioni di compilazione pulite e standard.
build_exe_options = {
    "packages": packages,
    "includes": ["matplotlib.backends.backend_tkagg"],
    "include_files": include_files,
    "include_msvcr": True, # Include le librerie C++ ridistribuibili
    "build_exe": build_dir_name
}

setup(
    name=f"{APP_NAME}App",
    version=APP_VERSION,
    description="Applicazione per la gestione e analisi di file.",
    options={"build_exe": build_exe_options},
    executables=[Executable(
        "winfile.py",
        base=base,
        target_name=f"{APP_NAME}.exe",
        icon="icona.ico"
    )]
)