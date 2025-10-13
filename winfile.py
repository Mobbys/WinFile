# winfile.py - v2.2
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
import threading
import requests
import webbrowser
from packaging import version
import tempfile # Aggiunto per la cartella temporanea
import subprocess # Aggiunto per avviare l'installer
import sys # Aggiunto per ottenere il percorso dell'eseguibile

# --- CONFIGURAZIONE AGGIORNAMENTI ---
GITHUB_REPO = "Mobbys/WinFile" 
CURRENT_VERSION = "4.7" # Questa deve corrispondere alla versione nel setup.py

# Il repository è pubblico, quindi non è necessario un token di accesso.

# --- Finestra delle Impostazioni ---
class SettingsWindow(ctk.CTkToplevel):
    # ... (il codice di questa classe rimane invariato) ...
    def __init__(self, master, config, save_callback):
        super().__init__(master)
        self.title("Impostazioni")
        self.geometry("400x250")
        self.transient(master)
        self.grab_set()
        self.config = config
        self.save_callback = save_callback
        self.appearance_mode_var = ctk.StringVar(value=self.config.get('theme', {}).get('appearance_mode', 'System'))
        self.color_theme_var = ctk.StringVar(value=self.config.get('theme', {}).get('color_theme', 'blue'))
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)
        appearance_label = ctk.CTkLabel(main_frame, text="Modalità Aspetto:")
        appearance_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        appearance_menu = ctk.CTkOptionMenu(main_frame, variable=self.appearance_mode_var, values=["Light", "Dark", "System"])
        appearance_menu.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        theme_label = ctk.CTkLabel(main_frame, text="Tema Colore:")
        theme_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        theme_menu = ctk.CTkOptionMenu(main_frame, variable=self.color_theme_var, values=["blue", "dark-blue", "green"])
        theme_menu.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        restart_label = ctk.CTkLabel(main_frame, text="È necessario riavviare l'applicazione per applicare le modifiche.", wraplength=350, text_color="gray")
        restart_label.grid(row=2, column=0, columnspan=2, padx=10, pady=15)
        save_button = ctk.CTkButton(main_frame, text="Salva e Chiudi", command=self.save_and_close)
        save_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
        main_frame.grid_columnconfigure(1, weight=1)

    def save_and_close(self):
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

        settings_button = ctk.CTkButton(top_bar, text="⚙️", font=ctk.CTkFont(size=20), width=32, height=32, command=self.open_settings_window)
        settings_button.pack(side="right", padx=10)

        # --- Creazione del Gestore di Schede (Tab View) ---
        self.tab_view = ctk.CTkTabview(self, anchor="nw")
        self.tab_view.pack(padx=10, pady=10, fill="both", expand=True)

        # --- Caricamento delle App dal file di configurazione ---
        self.load_apps()

        # --- Avvio controllo aggiornamenti in background ---
        update_thread = threading.Thread(target=self.check_for_updates, daemon=True)
        update_thread.start()

    def check_for_updates(self):
        """Contatta l'API di GitHub per verificare la presenza di una nuova release."""
        try:
            api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            
            latest_release = response.json()
            if not latest_release or latest_release.get('draft') or latest_release.get('prerelease'):
                return

            latest_version_str = latest_release['tag_name'].lstrip('v')
            
            if version.parse(latest_version_str) > version.parse(CURRENT_VERSION):
                self.master.after(0, self.start_silent_download, latest_release)

        except requests.exceptions.RequestException as e:
            print(f"Errore di rete durante il controllo degli aggiornamenti: {e}")
        except Exception as e:
            print(f"Errore generico durante il controllo degli aggiornamenti: {e}")

    def start_silent_download(self, release_info):
        """Avvia il download dell'aggiornamento in un thread separato."""
        assets = release_info.get('assets', [])
        
        download_url = None
        # NUOVA LOGICA: Diamo priorità al file .zip per l'aggiornamento automatico
        for asset in assets:
            if asset['name'].endswith('.zip'):
                download_url = asset['browser_download_url']
                break
        # Se non c'è lo zip, cerchiamo un installer come fallback
        if not download_url:
            for asset in assets:
                 if asset['name'].endswith(('.exe', '.msi')):
                    download_url = asset['browser_download_url']
                    break

        if download_url:
            print(f"Nuova versione trovata. Avvio download da: {download_url}")
            download_thread = threading.Thread(target=self.download_and_install, args=(download_url, release_info), daemon=True)
            download_thread.start()
        else:
            print("Aggiornamento trovato, ma nessun file .zip, .exe o .msi disponibile per il download.")

    def download_and_install(self, url, release_info):
        """
        Scarica il file di aggiornamento in una cartella temporanea 
        e poi chiede all'utente di installarlo.
        """
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()

            filename = url.split('/')[-1]
            temp_dir = tempfile.gettempdir()
            installer_path = os.path.join(temp_dir, filename)

            with open(installer_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            print(f"Download completato: {installer_path}")

            self.master.after(0, self.prompt_to_install, installer_path, release_info)

        except requests.exceptions.RequestException as e:
            print(f"Errore durante il download dell'aggiornamento: {e}")
        except Exception as e:
            print(f"Errore imprevisto durante il processo di aggiornamento: {e}")

    def prompt_to_install(self, installer_path, release_info):
        """
        Mostra un popup per informare che l'aggiornamento è pronto
        e chiede se si vuole procedere con l'installazione.
        """
        notes = "Nessuna nota di rilascio disponibile."
        if release_info and 'body' in release_info and release_info['body']:
            notes = release_info['body']

        message = (f"Un nuovo aggiornamento è stato scaricato ed è pronto per l'installazione!\n\n"
                   f"Novità:\n{notes}\n\n"
                   f"Vuoi chiudere l'applicazione e installarlo ora?")
        
        if messagebox.askyesno("Aggiornamento Pronto", message, parent=self):
            # NUOVA LOGICA: Controlla se è un aggiornamento portable (.zip) o un installer
            if installer_path.endswith('.zip'):
                self._launch_portable_updater(installer_path)
            else:
                try:
                    # Comportamento precedente per gli installer (.exe, .msi)
                    os.startfile(installer_path)
                    print("Avvio dell'installer e chiusura dell'applicazione...")
                    self.master.destroy()
                except Exception as e:
                    messagebox.showerror("Errore", f"Impossibile avviare l'installer.\n\nPercorso: {installer_path}\nErrore: {e}", parent=self)

    def _launch_portable_updater(self, zip_path):
        """
        Crea ed esegue uno script batch (.bat) di aggiornamento per gestire la
        sostituzione dei file della versione portable, senza richiedere Python.
        """
        try:
            app_path = sys.executable
            app_dir = os.path.dirname(app_path)
            
            if getattr(sys, 'frozen', False):
                restart_command = f'"{app_path}"'
            else:
                main_script_path = os.path.abspath(sys.argv[0])
                restart_command = f'"{app_path}" "{main_script_path}"'

            # --- CORREZIONE ---
            # Crea uno script PowerShell robusto che gestisce gli zip con e senza cartella radice
            # e sovrascrive correttamente i file esistenti usando Copy-Item.
            powershell_script = f"""
                $tempExtractPath = Join-Path -Path '{app_dir}' -ChildPath '_update_temp'
                if (Test-Path $tempExtractPath) {{ Remove-Item -Path $tempExtractPath -Recurse -Force }}
                New-Item -Path $tempExtractPath -ItemType Directory -Force | Out-Null
                
                Expand-Archive -Path '{zip_path}' -DestinationPath $tempExtractPath -Force
                
                $sourceFolder = $tempExtractPath
                $extractedItems = Get-ChildItem -Path $tempExtractPath
                if (($extractedItems.Count -eq 1) -and ($extractedItems[0].PSIsContainer)) {{
                    $sourceFolder = $extractedItems[0].FullName
                }}
                
                Copy-Item -Path ($sourceFolder + "\\*") -Destination '{app_dir}' -Recurse -Force
                
                Remove-Item -Path $tempExtractPath -Recurse -Force
            """
            
            # Codifica lo script in Base64 per passarlo a PowerShell in modo sicuro
            encoded_ps_script = base64.b64encode(powershell_script.encode('utf_16_le')).decode('utf-8')

            updater_script_code = f"""@echo off
title WinFile Updater
chcp 65001 > nul

set "ZIP_PATH={zip_path}"

rem Attende 4 secondi per permettere all'app principale di chiudersi
timeout /t 4 /nobreak > nul

rem Usa PowerShell per estrarre l'archivio in modo intelligente, gestendo la cartella radice
powershell -NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded_ps_script}

rem Pulisce il file zip scaricato
del "%ZIP_PATH%" > nul

rem Riavvia l'applicazione aggiornata
start "" {restart_command}

rem Il comando finale elimina lo script batch stesso
(goto) 2>nul & del "%~f0"
"""
            temp_dir = tempfile.gettempdir()
            updater_script_path = os.path.join(temp_dir, 'winfile_updater.bat')
            with open(updater_script_path, 'w', encoding='utf-8') as f:
                f.write(updater_script_code)

            subprocess.Popen([updater_script_path], creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP, close_fds=True)
            
            print("Avvio dello script di aggiornamento e chiusura dell'app principale...")
            self.master.destroy()

        except Exception as e:
            messagebox.showerror("Errore di Aggiornamento", f"Impossibile avviare il processo di aggiornamento automatico.\\n\\nErrore: {e}", parent=self)

    def open_settings_window(self):
        if self.settings_win is None or not self.settings_win.winfo_exists():
            self.settings_win = SettingsWindow(master=self, config=self.config, save_callback=self.save_config_func)
            self.settings_win.focus()
        else:
            self.settings_win.focus()

    def load_apps(self):
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
                else:
                    print(f"Attenzione: Il modulo '{module_name}' non ha una funzione 'create_tab'.")
            except Exception as e:
                print(f"Errore durante il caricamento del modulo '{module_name}': {e}")

# --- (Il resto del file rimane invariato) ---
class CTkRoot(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

app_instance = None

def handle_global_drop(event):
    global app_instance
    if not app_instance: return
    active_tab_name = app_instance.tab_view.get()
    if active_tab_name in app_instance.app_instances:
        target_app = app_instance.app_instances[active_tab_name]
        if hasattr(target_app, 'handle_drop') and callable(getattr(target_app, 'handle_drop')):
            target_app.handle_drop(event)

def save_config(config_data):
    try:
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
    except Exception as e:
        print(f"Errore: Impossibile salvare la configurazione in 'config.json': {e}")

if __name__ == "__main__":
    config = {
        "window": {"width": 1200, "height": 700, "min_width": 800, "min_height": 600},
        "theme": {"appearance_mode": "System", "color_theme": "blue"},
        "apps": [
            {"module": "app_liste_anteprime"}, 
            {"module": "app_controllo_immagini"},
            {"module": "app_controllo_pdf"}
        ]
    }
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pass

    theme_conf = config.get('theme', {})
    ctk.set_appearance_mode(theme_conf.get('appearance_mode', 'System')) 
    ctk.set_default_color_theme(theme_conf.get('color_theme', 'blue'))

    root = CTkRoot()
    root.title(f"WinFile v{CURRENT_VERSION}")
    
    win_conf = config.get('window', {})
    root.geometry(f"{win_conf.get('width', 1200)}x{win_conf.get('height', 700)}")
    root.minsize(win_conf.get('min_width', 800), win_conf.get('min_height', 600))

    app_instance = WinFileApp(master=root, config=config, save_config_func=save_config)
    app_instance.pack(fill="both", expand=True)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', handle_global_drop)

    root.mainloop()

