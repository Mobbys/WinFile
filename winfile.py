# winfile.py - v2.1
import customtkinter as ctk
import os
import importlib.util
from pathlib import Path
from tkinterdnd2 import DND_FILES, TkinterDnD
import json
from tkinter import messagebox
import base64
import io
from PIL import Image

# --- Dati Base64 per l'icona delle impostazioni (Grigio Medio Scuro) ---
GEAR_ICON_GRAY_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACx"
    "jwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAJnSURBVEhLzZVLaBNRFIb/Z2Z302wSjclg0qRJ"
    "a7RWo4J6URDEFy/iQ/xT8FN86EUPonjwopeCgiiIeFA8CF48CIIi6EU8CEHFQ3pSmjZN2qaZzO7O"
    "/H/czmSSTdP0Bw8cy+z85/zPzJk7gP/PsGg0gq7r8DwPgiCAIAjwer1+f3ieB5IkQVEUuK4LruvC"
    "NE2i0SjG4/E/hCRJgud5sCwLmqYBgGEYsCwLruvCNE2i0SgYjUbY7XbY7XbY7/fIZrMwm80wm80g"
    "CAJ83wfHcaBUKqGqquqqb2ZnZ/8CGI1GGI1GmM1mGI1GMAwDjuPg+/43ruvCNE183wfHcSCSyOjo"
    "KADwPA8syyIajQIAarUaHMfBYrGAIAjwPA9EIlE9k1ar/RewWCwAgCRJ4LouPM9DURTwer0wDAPL"
    "soDjODBNE6lUimze72x9C4VSCH3fhzAMuK4Lz/N8T9M0RCIRmEwmuK4LjuPg+/5bURQAgCRJ4Lou"
    "PM/DNE2IRCJwOp2w2+2w2+1wOp0wDAPf94PjOBCJRLquq6qqqv+L/wH/f2g5G0VpPAAAAABJRU5ErkJggg=="
)


# --- Finestra delle Impostazioni ---
class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, master, config, save_callback):
        super().__init__(master)

        self.title("Impostazioni")
        self.geometry("400x250")
        self.transient(master)  # Mantiene la finestra in primo piano
        self.grab_set()  # Blocca l'interazione con la finestra principale

        self.config = config
        self.save_callback = save_callback

        # Variabili per i widget
        self.appearance_mode_var = ctk.StringVar(value=self.config.get('theme', {}).get('appearance_mode', 'System'))
        self.color_theme_var = ctk.StringVar(value=self.config.get('theme', {}).get('color_theme', 'blue'))

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Opzione Modalità Aspetto
        appearance_label = ctk.CTkLabel(main_frame, text="Modalità Aspetto:")
        appearance_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        appearance_menu = ctk.CTkOptionMenu(main_frame, variable=self.appearance_mode_var, values=["Light", "Dark", "System"])
        appearance_menu.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # Opzione Tema Colore
        theme_label = ctk.CTkLabel(main_frame, text="Tema Colore:")
        theme_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        theme_menu = ctk.CTkOptionMenu(main_frame, variable=self.color_theme_var, values=["blue", "dark-blue", "green"])
        theme_menu.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # Messaggio di riavvio
        restart_label = ctk.CTkLabel(main_frame, text="È necessario riavviare l'applicazione per applicare le modifiche.", wraplength=350, text_color="gray")
        restart_label.grid(row=2, column=0, columnspan=2, padx=10, pady=15)

        # Pulsante Salva
        save_button = ctk.CTkButton(main_frame, text="Salva e Chiudi", command=self.save_and_close)
        save_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

        main_frame.grid_columnconfigure(1, weight=1)

    def save_and_close(self):
        """Aggiorna la configurazione, la salva su file e chiude la finestra."""
        if 'theme' not in self.config:
            self.config['theme'] = {}
        self.config['theme']['appearance_mode'] = self.appearance_mode_var.get()
        self.config['theme']['color_theme'] = self.color_theme_var.get()
        
        self.save_callback(self.config)
        
        messagebox.showinfo("Impostazioni Salvate", "Impostazioni salvate con successo. Riavvia l'applicazione per vedere le modifiche.", parent=self)
        
        self.destroy()

# --- Classe Principale dell'Applicazione ---
class WinFileApp(ctk.CTkFrame):
    def __init__(self, master, config, save_config_func):
        super().__init__(master)
        
        self.config = config
        self.save_config_func = save_config_func
        self.app_instances = {} 
        self.settings_win = None

        # --- Barra Superiore per Titolo e Impostazioni ---
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=10, pady=(5, 0))
        
        title_label = ctk.CTkLabel(top_bar, text="WinFile", font=ctk.CTkFont(size=16, weight="bold"))
        title_label.pack(side="left", padx=10)

        # Carica l'icona e crea il pulsante con uno sfondo visibile
        try:
            icon_data_gray = base64.b64decode(GEAR_ICON_GRAY_BASE64)
            icon_image = Image.open(io.BytesIO(icon_data_gray))

            self.settings_icon = ctk.CTkImage(
                light_image=icon_image,
                dark_image=icon_image,
                size=(24, 24)
            )
            # Crea un pulsante standard che avrà sempre uno sfondo visibile
            settings_button = ctk.CTkButton(top_bar, 
                                            text="", 
                                            image=self.settings_icon,
                                            width=28, # Mantiene una dimensione fissa
                                            height=28,
                                            command=self.open_settings_window)

        except Exception as e:
            print(f"Errore caricamento icona: {e}. Uso un pulsante di testo.")
            # Pulsante di fallback, ora anche questo con uno sfondo visibile
            settings_button = ctk.CTkButton(top_bar, 
                                            text="⚙️",
                                            font=ctk.CTkFont(size=20),
                                            width=32,
                                            height=32,
                                            command=self.open_settings_window)
        
        settings_button.pack(side="right", padx=10)

        # --- Creazione del Gestore di Schede (Tab View) ---
        self.tab_view = ctk.CTkTabview(self, anchor="nw")
        self.tab_view.pack(padx=10, pady=10, fill="both", expand=True)

        # --- Caricamento delle App dal file di configurazione ---
        self.load_apps()

    def open_settings_window(self):
        """Apre la finestra delle impostazioni, assicurandosi che ne esista solo una."""
        if self.settings_win is None or not self.settings_win.winfo_exists():
            self.settings_win = SettingsWindow(master=self, config=self.config, save_callback=self.save_config_func)
            self.settings_win.focus()
        else:
            self.settings_win.focus()

    def load_apps(self):
        """Carica le app specificate nel file di configurazione, nell'ordine dato."""
        apps_path = Path("apps")
        if 'apps' not in self.config or not isinstance(self.config['apps'], list):
            print("Errore: La configurazione delle app non è valida o mancante in config.json.")
            return
        for app_config in self.config['apps']:
            module_name = app_config.get('module')
            if not module_name:
                print(f"Attenzione: Trovata una voce non valida nella configurazione delle app: {app_config}")
                continue
            file_path = apps_path / f"{module_name}.py"
            if not file_path.exists():
                print(f"Attenzione: Il file del modulo '{module_name}' non è stato trovato in '{file_path}'.")
                continue
            try:
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                app_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(app_module)
                if hasattr(app_module, "create_tab"):
                    tab_name, instance = app_module.create_tab(self.tab_view)
                    if tab_name and instance:
                        self.app_instances[tab_name] = instance
                        print(f"Modulo '{module_name}' caricato con successo nella scheda '{tab_name}'.")
                else:
                    print(f"Attenzione: Il modulo '{module_name}' non ha una funzione 'create_tab'.")
            except Exception as e:
                print(f"Errore durante il caricamento del modulo '{module_name}': {e}")

# --- Classe Personalizzata per la Finestra Principale ---
class CTkRoot(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

# --- Variabile globale per accedere all'istanza dell'app principale ---
app_instance = None

def handle_global_drop(event):
    """Gestisce il drag and drop globale e lo inoltra alla scheda attiva."""
    global app_instance
    if not app_instance: return
    active_tab_name = app_instance.tab_view.get()
    if active_tab_name in app_instance.app_instances:
        target_app = app_instance.app_instances[active_tab_name]
        if hasattr(target_app, 'handle_drop') and callable(getattr(target_app, 'handle_drop')):
            target_app.handle_drop(event)

def save_config(config_data):
    """Salva il dizionario di configurazione nel file config.json."""
    try:
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
        print("Info: Configurazione salvata con successo in 'config.json'.")
    except Exception as e:
        print(f"Errore: Impossibile salvare la configurazione in 'config.json': {e}")

# --- Avvio dell'Applicazione ---
if __name__ == "__main__":
    config = {
        "window": {"width": 1200, "height": 700, "min_width": 800, "min_height": 600},
        "theme": {"appearance_mode": "System", "color_theme": "blue"},
        "apps": [{"module": "app_liste_anteprime"}, {"module": "app_controllo_immagini"}]
    }
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Info: File 'config.json' non trovato o non valido. Uso la configurazione di default. Dettagli: {e}")
        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
                print("Info: Creato un file 'config.json' di default.")
        except Exception as write_e:
            print(f"Errore: Impossibile creare il file 'config.json' di default: {write_e}")

    theme_conf = config.get('theme', {})
    ctk.set_appearance_mode(theme_conf.get('appearance_mode', 'System')) 
    ctk.set_default_color_theme(theme_conf.get('color_theme', 'blue'))

    root = CTkRoot()
    root.title("WinFile v2.1")
    
    win_conf = config.get('window', {})
    root.geometry(f"{win_conf.get('width', 1200)}x{win_conf.get('height', 700)}")
    root.minsize(win_conf.get('min_width', 800), win_conf.get('min_height', 600))

    app_instance = WinFileApp(master=root, config=config, save_config_func=save_config)
    app_instance.pack(fill="both", expand=True)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', handle_global_drop)

    root.mainloop()
