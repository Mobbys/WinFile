import customtkinter as ctk

def create_tab(tab_view):
    """
    Questa è la funzione standard che 'winfile.py' cerca e chiama
    per costruire l'interfaccia di questa specifica app.

    Args:
        tab_view: L'oggetto CTkTabview dalla finestra principale a cui aggiungere la scheda.
    """
    # Aggiunge una nuova scheda al TabView principale. Il nome della scheda sarà "Blocco Note".
    tab = tab_view.add("Blocco Note")

    # --- Inizio del Codice Specifico dell'App Blocco Note ---
    # Tutti i widget di questa app devono essere inseriti dentro 'tab'.

    # Creiamo un frame per contenere i pulsanti
    button_frame = ctk.CTkFrame(master=tab, fg_color="transparent")
    button_frame.pack(padx=10, pady=10, fill="x")

    # Un'etichetta descrittiva
    label = ctk.CTkLabel(master=button_frame, text="Un semplice blocco note")
    label.pack(side="left", padx=(0, 20))

    # Un pulsante di esempio (per ora non fa nulla)
    save_button = ctk.CTkButton(master=button_frame, text="Salva (Esempio)")
    save_button.pack(side="right")

    # L'area di testo principale
    textbox = ctk.CTkTextbox(master=tab, corner_radius=8, font=("Arial", 14))
    textbox.pack(padx=10, pady=(0, 10), fill="both", expand=True)
    
    textbox.insert("0.0", "Scrivi qualcosa qui...")
    # --- Fine del Codice Specifico dell'App ---
