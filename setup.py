import sys
from cx_Freeze import setup, Executable

# --- Dati dell'applicazione (modifica qui per versioni future) ---
APP_NAME = "WinFile"
APP_VERSION = "3.0"

# --- Opzioni di Base ---
packages = ["customtkinter", "tkinterdnd2", "PIL", "fitz", "reportlab", "requests", "packaging"]

# --- Gestione dei File Esterni ---
include_files = [
    ('apps', 'apps'),
    ('config.json', 'config.json')
]

# --- Configurazione per l'eseguibile ---
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# --- Costruisci il nome della cartella di output ---
# Questa riga crea una stringa tipo "build/WinFile 2.5"
# cx_Freeze crea la cartella 'build' di default, noi specifichiamo la sottocartella
build_dir_name = f"build/{APP_NAME} {APP_VERSION}"

setup(
    name=f"{APP_NAME}App",
    version=APP_VERSION,
    description="Applicazione per la gestione e analisi di file.",
    options={"build_exe": {
        'packages': packages,
        'include_files': include_files,
        'include_msvcr': True,
        # CORREZIONE: L'opzione corretta per la cartella di output è 'build_exe'
        'build_exe': build_dir_name,
    }},
    executables=[Executable(
        "winfile.py",
        base=base,
        target_name=f"{APP_NAME}.exe",
        icon="icona.ico"
    )]
)
