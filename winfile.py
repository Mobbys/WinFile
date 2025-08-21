# winfile.py - v1.4
import customtkinter as ctk
import os
import importlib.util
from pathlib import Path
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- Classe Personalizzata per la Finestra Principale ---
# Questa classe eredita sia da ctk.CTk (per l'interfaccia moderna) sia da
# TkinterDnD.DnDWrapper (per aggiungere la funzionalità di drag and drop).
# Questo approccio risolve le incompatibilità di rendering tra le due librerie,
# prevenendo l'effetto "finestra trasparente" durante le operazioni bloccanti.
class CTkRoot(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Questa linea è cruciale per inizializzare il sistema DnD sulla nuova finestra
        self.TkdndVersion = TkinterDnD._require(self)

# --- Variabile globale per accedere all'istanza dell'app principale ---
app_instance = None

class WinFileApp(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        
        # Dizionario per memorizzare le istanze delle app caricate in ogni scheda
        self.app_instances = {} 

        # --- Creazione del Gestore di Schede (Tab View) ---
        self.tab_view = ctk.CTkTabview(self, anchor="nw")
        self.tab_view.pack(padx=10, pady=10, fill="both", expand=True)

        # --- Caricamento Dinamico delle App dalla cartella 'apps' ---
        self.load_apps()

    def load_apps(self):
        """
        Scandaglia la cartella 'apps', importa ogni modulo python che inizia
        con 'app_' e chiama la sua funzione 'create_tab' per inizializzarlo.
        """
        apps_path = Path("apps")
        
        if not apps_path.is_dir():
            print(f"Errore: La cartella '{apps_path}' non è stata trovata.")
            return

        for file_path in apps_path.glob("app_*.py"):
            module_name = file_path.stem
            
            try:
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                app_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(app_module)

                if hasattr(app_module, "create_tab"):
                    # La funzione create_tab restituisce il nome della scheda e l'istanza dell'app
                    tab_name, instance = app_module.create_tab(self.tab_view)
                    if tab_name and instance:
                        # Memorizziamo l'istanza usando il nome della scheda come chiave
                        # per poterla ritrovare in seguito (es. per il drag and drop)
                        self.app_instances[tab_name] = instance
                        print(f"Modulo '{module_name}' caricato con successo nella scheda '{tab_name}'.")
                else:
                    print(f"Attenzione: Il modulo '{module_name}' non ha una funzione 'create_tab'.")

            except Exception as e:
                print(f"Errore durante il caricamento del modulo '{module_name}': {e}")

def handle_global_drop(event):
    """
    Questa funzione viene chiamata quando un file viene trascinato sulla finestra.
    Controlla quale scheda è attiva e inoltra l'evento all'istanza corretta.
    """
    global app_instance
    if not app_instance:
        return

    # Ottiene il nome della scheda attualmente visualizzata
    active_tab_name = app_instance.tab_view.get()
    
    # Controlla se per questa scheda esiste un'istanza di app registrata
    if active_tab_name in app_instance.app_instances:
        target_app = app_instance.app_instances[active_tab_name]
        
        # Se l'istanza dell'app ha un metodo 'handle_drop', lo chiamiamo,
        # passandogli l'evento del rilascio.
        if hasattr(target_app, 'handle_drop') and callable(getattr(target_app, 'handle_drop')):
            target_app.handle_drop(event)

# --- Avvio dell'Applicazione ---
if __name__ == "__main__":
    # Usiamo la nostra nuova classe CTkRoot al posto di TkinterDnD.Tk()
    root = CTkRoot()
    
    # La patch di compatibilità precedente non è più necessaria perché la nostra
    # finestra principale è ora un oggetto CTk a tutti gli effetti.

    root.title("WinFile v1.4")
    root.geometry("1200x700")
    root.minsize(800, 600)

    ctk.set_appearance_mode("System") 
    ctk.set_default_color_theme("blue")

    app_instance = WinFileApp(master=root)
    app_instance.pack(fill="both", expand=True)

    # Registra l'intera finestra come area di rilascio per i file
    root.drop_target_register(DND_FILES)
    # Associa l'evento di rilascio (<<Drop>>) alla nostra funzione di gestione
    root.dnd_bind('<<Drop>>', handle_global_drop)

    root.mainloop()
