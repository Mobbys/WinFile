# app_controllo_immagini.py - v2.0
import customtkinter as ctk
import os
import threading
from PIL import Image, ExifTags
import math

# Disabilita il limite di dimensione per le immagini grandi
Image.MAX_IMAGE_PIXELS = None

class ImageCheckerApp(ctk.CTkFrame):
    """
    Applicazione per visualizzare una o più immagini e le loro proprietà principali.
    """
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        self.loaded_images = []
        self.thumbnail_buttons = []
        self.active_image_index = -1
        self.ctk_image = None
        self.image_width_px = None
        self.image_height_px = None
        self._is_updating_dimensions = False

        # --- Layout Principale ---
        self.grid_columnconfigure(0, weight=0, minsize=120) # Colonna miniature (stretta, fissa)
        self.grid_columnconfigure(1, weight=1)              # Colonna contenuto principale (espandibile)
        self.grid_rowconfigure(0, weight=1)

        # --- Frame Lista Miniature (Sinistra) ---
        self.thumbnail_list_frame = ctk.CTkScrollableFrame(self, label_text="Immagini")
        self.thumbnail_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- Frame Contenuto Principale (Destra) ---
        main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_content_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        main_content_frame.grid_columnconfigure(0, weight=1)
        main_content_frame.grid_rowconfigure(0, weight=0) # Info in alto
        main_content_frame.grid_rowconfigure(1, weight=1) # Contenuto sotto

        # --- Frame Info Nome/Percorso (Destra, in Cima) ---
        top_info_frame = ctk.CTkFrame(main_content_frame)
        top_info_frame.grid(row=0, column=0, sticky="ew")
        top_info_frame.grid_columnconfigure(1, weight=1)
        
        self.info_vars = {}
        top_prop_labels = ["Nome file:", "Percorso:"]
        for i, label_text in enumerate(top_prop_labels):
            label = ctk.CTkLabel(top_info_frame, text=label_text, anchor="e", font=ctk.CTkFont(weight="bold"))
            label.grid(row=i, column=0, padx=10, pady=2, sticky="ne")
            self.info_vars[label_text] = ctk.StringVar(value="-")
            value_label = ctk.CTkLabel(top_info_frame, textvariable=self.info_vars[label_text], anchor="w", wraplength=600)
            value_label.grid(row=i, column=1, padx=10, pady=2, sticky="w")

        # --- Frame Contenuto Inferiore (Dettagli a sx, Anteprima a dx) ---
        bottom_content_frame = ctk.CTkFrame(main_content_frame, fg_color="transparent")
        bottom_content_frame.grid(row=1, column=0, pady=(10, 0), sticky="nsew")
        bottom_content_frame.grid_columnconfigure(0, weight=1, minsize=320) # Dettagli
        bottom_content_frame.grid_columnconfigure(1, weight=4)             # Anteprima (molto più grande)
        bottom_content_frame.grid_rowconfigure(0, weight=1)

        # --- Frame Dettagli (Sinistra del contenuto inferiore) ---
        details_frame = ctk.CTkFrame(bottom_content_frame)
        details_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        details_frame.grid_columnconfigure(1, weight=1)

        # --- Frame Anteprima (Destra del contenuto inferiore) ---
        preview_frame = ctk.CTkFrame(bottom_content_frame)
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.image_label = ctk.CTkLabel(preview_frame, text="", text_color="gray")
        self.image_label.grid(row=0, column=0, sticky="nsew")

        # Popola il frame dei dettagli
        detail_prop_labels = ["Dimensioni:", "Modo colore:", "Dimensione file:", "Rapporto:"]
        row_index = 0
        for label_text in detail_prop_labels:
            label = ctk.CTkLabel(details_frame, text=label_text, anchor="e", font=ctk.CTkFont(weight="bold"))
            label.grid(row=row_index, column=0, padx=10, pady=2, sticky="e")
            self.info_vars[label_text] = ctk.StringVar(value="-")
            value_label = ctk.CTkLabel(details_frame, textvariable=self.info_vars[label_text], anchor="w")
            value_label.grid(row=row_index, column=1, padx=10, pady=2, sticky="w")
            row_index += 1
        
        # Campi DPI, Larghezza, Altezza
        self.dpi_var = ctk.StringVar()
        self.width_cm_var = ctk.StringVar()
        self.height_cm_var = ctk.StringVar()
        
        dpi_label = ctk.CTkLabel(details_frame, text="Risoluzione (DPI):", anchor="e", font=ctk.CTkFont(weight="bold"))
        dpi_label.grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
        self.dpi_entry = ctk.CTkEntry(details_frame, textvariable=self.dpi_var)
        self.dpi_entry.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")
        row_index += 1

        width_label = ctk.CTkLabel(details_frame, text="Larghezza (cm):", anchor="e", font=ctk.CTkFont(weight="bold"))
        width_label.grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
        self.width_cm_entry = ctk.CTkEntry(details_frame, textvariable=self.width_cm_var)
        self.width_cm_entry.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")
        row_index += 1

        height_label = ctk.CTkLabel(details_frame, text="Altezza (cm):", anchor="e", font=ctk.CTkFont(weight="bold"))
        height_label.grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
        self.height_cm_entry = ctk.CTkEntry(details_frame, textvariable=self.height_cm_var)
        self.height_cm_entry.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")
        row_index += 1
        
        self.width_cm_var.trace_add("write", self._update_from_width)
        self.height_cm_var.trace_add("write", self._update_from_height)
        self.dpi_var.trace_add("write", self._update_from_dpi)

        # Campo Distanza di Visione
        self.distance_var = ctk.StringVar(value="-")
        distance_label = ctk.CTkLabel(details_frame, text="Distanza visione:", anchor="e", font=ctk.CTkFont(weight="bold"))
        distance_label.grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
        distance_value_label = ctk.CTkLabel(details_frame, textvariable=self.distance_var, anchor="w")
        distance_value_label.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

    def handle_drop(self, event):
        """Gestisce il rilascio di uno o più file."""
        try:
            paths = self.tk.splitlist(event.data)
            if paths:
                self.loaded_images.clear()
                for path in paths:
                    info = self._process_image(path)
                    if info:
                        self.loaded_images.append(info)
                self.after(0, self._populate_thumbnail_list)
        except Exception as e:
            print(f"Errore durante il drop: {e}")

    def _process_image(self, image_path):
        """Analizza una singola immagine e restituisce un dizionario con i suoi dati."""
        try:
            with Image.open(image_path) as img:
                if hasattr(img, '_getexif'):
                    exif = img._getexif()
                    if exif:
                        orientation_key = next((key for key, value in ExifTags.TAGS.items() if value == 'Orientation'), None)
                        if orientation_key in exif:
                            orientation = exif[orientation_key]
                            if orientation == 3: img = img.rotate(180, expand=True)
                            elif orientation == 6: img = img.rotate(270, expand=True)
                            elif orientation == 8: img = img.rotate(90, expand=True)
                
                width_px, height_px = img.size
                gcd = math.gcd(width_px, height_px)
                
                file_size = os.path.getsize(image_path)
                if file_size < 1024: size_str = f"{file_size} Bytes"
                elif file_size < 1024**2: size_str = f"{file_size/1024:.2f} KB"
                else: size_str = f"{file_size/1024**2:.2f} MB"

                return {
                    "path": image_path,
                    "image_obj": img.copy(),
                    "width_px": width_px,
                    "height_px": height_px,
                    "Nome file:": os.path.basename(image_path),
                    "Percorso:": os.path.dirname(image_path),
                    "Dimensioni:": f"{width_px} x {height_px} px",
                    "Modo colore:": img.mode,
                    "Dimensione file:": size_str,
                    "Rapporto:": f"{width_px//gcd}:{height_px//gcd}",
                    "Risoluzione:": img.info.get('dpi', (96, 96))[0]
                }
        except Exception as e:
            print(f"Errore nell'analisi dell'immagine {image_path}: {e}")
            return None

    def _populate_thumbnail_list(self):
        """Crea la lista di miniature a sinistra."""
        for widget in self.thumbnail_list_frame.winfo_children():
            widget.destroy()
        self.thumbnail_buttons.clear()

        for index, img_data in enumerate(self.loaded_images):
            img_copy = img_data["image_obj"].copy()
            img_copy.thumbnail((80, 80), Image.Resampling.LANCZOS) # Miniatura più grande
            thumb_img = ctk.CTkImage(light_image=img_copy, dark_image=img_copy, size=img_copy.size)
            
            btn = ctk.CTkButton(self.thumbnail_list_frame, 
                                image=thumb_img, 
                                text=img_data["Nome file:"], 
                                compound="top",
                                font=ctk.CTkFont(size=11), # Font leggibile
                                command=lambda idx=index: self._display_image_details(idx))
            btn.pack(padx=5, pady=5, fill="x")
            self.thumbnail_buttons.append(btn)
        
        if self.loaded_images:
            self._display_image_details(0)
        else:
            self._clear_details()

    def _display_image_details(self, index):
        """Mostra i dettagli e l'anteprima dell'immagine selezionata."""
        self.active_image_index = index
        image_info = self.loaded_images[index]

        # Evidenzia il pulsante della miniatura selezionata
        for i, btn in enumerate(self.thumbnail_buttons):
            if i == index:
                btn.configure(fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"])
            else:
                btn.configure(fg_color="transparent")

        self.image_width_px = image_info["width_px"]
        self.image_height_px = image_info["height_px"]

        self._is_updating_dimensions = True
        for key, var in self.info_vars.items():
            var.set(image_info.get(key, "-"))
        
        initial_dpi = image_info.get("Risoluzione:", 96)
        self.dpi_var.set(str(int(initial_dpi)))
        self._is_updating_dimensions = False
        
        # Aggiorna l'anteprima grande
        img_obj = image_info["image_obj"]
        container_width = self.image_label.winfo_width()
        container_height = self.image_label.winfo_height()
        img_copy = img_obj.copy()
        img_copy.thumbnail((container_width - 20, container_height - 20), Image.Resampling.LANCZOS)
        self.ctk_image = ctk.CTkImage(light_image=img_copy, dark_image=img_copy, size=img_copy.size)
        self.image_label.configure(image=self.ctk_image, text="")

        # Forza l'aggiornamento dei calcoli basati sui DPI iniziali
        self._update_from_dpi()

    def _clear_details(self):
        """Pulisce i campi dei dettagli quando non ci sono immagini."""
        self._is_updating_dimensions = True
        for var in self.info_vars.values():
            var.set("-")
        self.dpi_var.set("")
        self.width_cm_var.set("")
        self.height_cm_var.set("")
        self.distance_var.set("-")
        self._is_updating_dimensions = False
        self.image_label.configure(image=None, text="Trascina una o più immagini qui")
        self.image_width_px = None
            
    def _calculate_and_update_all(self, source, value):
        """Funzione centralizzata per ricalcolare tutti i valori dipendenti."""
        if self._is_updating_dimensions or not self.image_width_px:
            return

        self._is_updating_dimensions = True
        try:
            if source == "dpi":
                new_dpi = int(value)
                if new_dpi <= 0: raise ValueError
                new_width_cm = (self.image_width_px / new_dpi) * 2.54
                new_height_cm = (self.image_height_px / new_dpi) * 2.54
            elif source == "width":
                new_width_cm = float(value)
                if new_width_cm <= 0: raise ValueError
                aspect_ratio = self.image_height_px / self.image_width_px
                new_height_cm = new_width_cm * aspect_ratio
                new_dpi = (self.image_width_px / new_width_cm) * 2.54
            elif source == "height":
                new_height_cm = float(value)
                if new_height_cm <= 0: raise ValueError
                aspect_ratio = self.image_width_px / self.image_height_px
                new_width_cm = new_height_cm * aspect_ratio
                new_dpi = (self.image_height_px / new_height_cm) * 2.54

            distance_m = 80 / new_dpi

            # Aggiorna tutte le variabili
            if source != "dpi": self.dpi_var.set(f"{int(new_dpi)}")
            if source != "width": self.width_cm_var.set(f"{new_width_cm:.2f}")
            if source != "height": self.height_cm_var.set(f"{new_height_cm:.2f}")
            self.distance_var.set(f"~ {distance_m:.2f} m")

        except (ValueError, ZeroDivisionError):
            pass # Ignora input non validi o calcoli impossibili
        finally:
            self._is_updating_dimensions = False

    def _update_from_width(self, *args):
        self._calculate_and_update_all("width", self.width_cm_var.get().replace(",", "."))

    def _update_from_height(self, *args):
        self._calculate_and_update_all("height", self.height_cm_var.get().replace(",", "."))

    def _update_from_dpi(self, *args):
        self._calculate_and_update_all("dpi", self.dpi_var.get())


# --- Funzione di caricamento richiesta da winfile.py ---
def create_tab(tab_view):
    """
    Crea la scheda "Controllo immagini", inizializza l'app e restituisce
    il nome della scheda e l'istanza dell'app per la gestione eventi.
    """
    tab_name = "Controllo immagini"
    tab = tab_view.add(tab_name)
    
    app_instance = ImageCheckerApp(master=tab)
    
    # Restituisce il nome della scheda e l'istanza dell'app
    return tab_name, app_instance
