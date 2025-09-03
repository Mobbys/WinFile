# apps/app_controllo_pdf.py - v16.4 (Semplificazione UI Ridimensionamento)
import customtkinter as ctk
import os
import fitz  # PyMuPDF
from tkinter import messagebox, filedialog, colorchooser
from PIL import Image
import traceback
import re

# Prova a importare la libreria per il Drag & Drop
try:
    from tkinterdnd2 import DND_FILES
    DND_SUPPORT = True
except ImportError:
    DND_SUPPORT = False

# Disabilita il limite di dimensione per le immagini grandi
Image.MAX_IMAGE_PIXELS = None

# Dizionario con formati pagina standard in punti (1 punto = 1/72 di pollice)
PAGE_FORMATS_PT = {
    "A3 (29.7x42.0 cm)": (842, 1191),
    "A4 (21.0x29.7 cm)": (595, 842),
    "A5 (14.8x21.0 cm)": (420, 595),
}
MM_TO_PT = 2.83465
CM_TO_PT = 28.3465

class PageSelectionDialog(ctk.CTkToplevel):
    """
    Una finestra di dialogo per selezionare pagine specifiche da un documento PDF,
    mostrando le miniature.
    """
    def __init__(self, master, doc, title="Seleziona Pagine", button_text="Aggiungi"):
        super().__init__(master)
        self.doc = doc
        self.selected_pages = []

        self.title(title)
        self.geometry("500x600")
        self.transient(master)
        self.grab_set()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        scroll_frame = ctk.CTkScrollableFrame(self, label_text=f"Pagine in '{os.path.basename(self.doc.name)}'")
        scroll_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.checkbox_vars = []
        self.thumbnail_images = []

        for i, page in enumerate(self.doc):
            row_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=5, pady=5)
            row_frame.grid_columnconfigure(1, weight=1)

            thumb_img = self._get_page_thumbnail(page)
            self.thumbnail_images.append(thumb_img)
            
            thumb_label = ctk.CTkLabel(row_frame, image=thumb_img, text="")
            thumb_label.grid(row=0, column=0, padx=(5,10))

            var = ctk.StringVar(value="off")
            cb = ctk.CTkCheckBox(row_frame, text=f"Pagina {i + 1}", variable=var, onvalue="on", offvalue="off")
            cb.grid(row=0, column=1, sticky="w")
            self.checkbox_vars.append(var)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        button_frame.grid_columnconfigure((0,1,2,3), weight=1)
        
        select_all_btn = ctk.CTkButton(button_frame, text="Seleziona Tutti", command=self.select_all)
        select_all_btn.grid(row=0, column=0, padx=5)

        deselect_all_btn = ctk.CTkButton(button_frame, text="Deseleziona Tutti", command=self.deselect_all)
        deselect_all_btn.grid(row=0, column=1, padx=5)

        confirm_btn = ctk.CTkButton(button_frame, text=button_text, command=self.confirm_selection, fg_color="green", hover_color="darkgreen")
        confirm_btn.grid(row=0, column=2, padx=5)

        cancel_btn = ctk.CTkButton(button_frame, text="Annulla", command=self.destroy, fg_color="gray", hover_color="darkgray")
        cancel_btn.grid(row=0, column=3, padx=5)

    def _get_page_thumbnail(self, page, size=(80, 113)):
        pix = page.get_pixmap(dpi=72)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail(size, Image.Resampling.LANCZOS)
        return ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

    def select_all(self):
        for var in self.checkbox_vars:
            var.set("on")

    def deselect_all(self):
        for var in self.checkbox_vars:
            var.set("off")

    def confirm_selection(self):
        self.selected_pages = [i for i, var in enumerate(self.checkbox_vars) if var.get() == "on"]
        self.destroy()

class AddMarginsDialog(ctk.CTkToplevel):
    """Dialogo per aggiungere margini a una o più pagine."""
    def __init__(self, master):
        super().__init__(master)
        self.result = None
        self.title("Aggiungi Margini")
        self.geometry("400x300")
        self.transient(master)
        self.grab_set()

        self.grid_columnconfigure(1, weight=1)

        # Margini
        margins_frame = ctk.CTkFrame(self, fg_color="transparent")
        margins_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        margins_frame.grid_columnconfigure((1, 3), weight=1)

        self.margin_top_var = ctk.StringVar(value="10")
        self.margin_bottom_var = ctk.StringVar(value="10")
        self.margin_left_var = ctk.StringVar(value="10")
        self.margin_right_var = ctk.StringVar(value="10")

        ctk.CTkLabel(margins_frame, text="Sopra (mm):").grid(row=0, column=0, padx=(0,5), pady=5, sticky="w")
        ctk.CTkEntry(margins_frame, textvariable=self.margin_top_var).grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(margins_frame, text="Sotto (mm):").grid(row=1, column=0, padx=(0,5), pady=5, sticky="w")
        ctk.CTkEntry(margins_frame, textvariable=self.margin_bottom_var).grid(row=1, column=1, sticky="ew")
        ctk.CTkLabel(margins_frame, text="Sinistra (mm):").grid(row=0, column=2, padx=(10,5), pady=5, sticky="w")
        ctk.CTkEntry(margins_frame, textvariable=self.margin_left_var).grid(row=0, column=3, sticky="ew")
        ctk.CTkLabel(margins_frame, text="Destra (mm):").grid(row=1, column=2, padx=(10,5), pady=5, sticky="w")
        ctk.CTkEntry(margins_frame, textvariable=self.margin_right_var).grid(row=1, column=3, sticky="ew")


        # Colore
        color_label = ctk.CTkLabel(self, text="Colore Margini:")
        color_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.color_hex = "#FFFFFF"
        self.color_button = ctk.CTkButton(self, text="Seleziona Colore", command=self.choose_color)
        self.color_button.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.color_button.configure(fg_color=self.color_hex, text_color=self._get_contrasting_text_color(self.color_hex))

        # Ambito di applicazione
        scope_label = ctk.CTkLabel(self, text="Applica a:")
        scope_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.scope_var = ctk.StringVar(value="current")
        
        radio_frame = ctk.CTkFrame(self, fg_color="transparent")
        radio_frame.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkRadioButton(radio_frame, text="Pagina Corrente", variable=self.scope_var, value="current").pack(anchor="w")
        ctk.CTkRadioButton(radio_frame, text="Tutte le Pagine", variable=self.scope_var, value="all").pack(anchor="w")
        ctk.CTkRadioButton(radio_frame, text="Pagine Selezionate", variable=self.scope_var, value="selected").pack(anchor="w")

        # Pulsanti di conferma
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        apply_btn = ctk.CTkButton(button_frame, text="Applica", command=self.apply, fg_color="green")
        apply_btn.pack(side="left", padx=10)
        
        cancel_btn = ctk.CTkButton(button_frame, text="Annulla", command=self.destroy, fg_color="gray")
        cancel_btn.pack(side="left", padx=10)

    def choose_color(self):
        color_code = colorchooser.askcolor(title="Scegli un colore", parent=self)
        if color_code and color_code[1]:
            self.color_hex = color_code[1]
            self.color_button.configure(fg_color=self.color_hex, text_color=self._get_contrasting_text_color(self.color_hex))

    def _get_contrasting_text_color(self, hex_color):
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        brightness = (r * 299 + g * 587 + b * 114) / 1000
        return "black" if brightness > 125 else "white"

    def apply(self):
        try:
            self.result = {
                "margin_top": float(self.margin_top_var.get()),
                "margin_bottom": float(self.margin_bottom_var.get()),
                "margin_left": float(self.margin_left_var.get()),
                "margin_right": float(self.margin_right_var.get()),
                "color_hex": self.color_hex,
                "scope": self.scope_var.get()
            }
            self.destroy()
        except ValueError:
            messagebox.showerror("Valore non valido", "Inserisci un valore numerico per i margini.", parent=self)

class ResizePageDialog(ctk.CTkToplevel):
    """Dialogo per ridimensionare le pagine a una dimensione specifica."""
    def __init__(self, master, current_page=None):
        super().__init__(master)
        self.result = None
        self.title("Ridimensiona Pagine")
        self.geometry("450x330")
        self.transient(master)
        self.grab_set()

        self.current_page = current_page
        self.aspect_ratio = None
        self._is_updating_dims = False

        if self.current_page:
            page_rect = self.current_page.rect
            page_width_pt = page_rect.width
            page_height_pt = page_rect.height
            if page_width_pt > 0 and page_height_pt > 0:
                self.aspect_ratio = page_width_pt / page_height_pt
            initial_width_cm = f"{page_width_pt / CM_TO_PT:.2f}"
            initial_height_cm = f"{page_height_pt / CM_TO_PT:.2f}"
        else:
            a4_w_pt, a4_h_pt = PAGE_FORMATS_PT["A4 (21.0x29.7 cm)"]
            initial_width_cm = f"{a4_w_pt / CM_TO_PT:.2f}"
            initial_height_cm = f"{a4_h_pt / CM_TO_PT:.2f}"

        self.grid_columnconfigure(1, weight=1)

        # Dimensioni Personalizzate
        size_frame = ctk.CTkFrame(self, fg_color="transparent")
        size_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        size_frame.grid_columnconfigure((1, 3), weight=1)

        self.width_var = ctk.StringVar(value=initial_width_cm)
        self.height_var = ctk.StringVar(value=initial_height_cm)

        ctk.CTkLabel(size_frame, text="Larghezza (cm):").grid(row=0, column=0, padx=(0, 5), pady=5, sticky="w")
        self.width_entry = ctk.CTkEntry(size_frame, textvariable=self.width_var)
        self.width_entry.grid(row=0, column=1, sticky="ew")
        
        ctk.CTkLabel(size_frame, text="Altezza (cm):").grid(row=0, column=2, padx=(10, 5), pady=5, sticky="w")
        self.height_entry = ctk.CTkEntry(size_frame, textvariable=self.height_var)
        self.height_entry.grid(row=0, column=3, sticky="ew")

        # Mantieni Proporzioni
        self.aspect_ratio_var = ctk.BooleanVar(value=True)
        self.aspect_ratio_checkbox = ctk.CTkCheckBox(self, text="Mantieni proporzioni", variable=self.aspect_ratio_var)
        self.aspect_ratio_checkbox.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        
        if self.aspect_ratio is None:
            self.aspect_ratio_checkbox.configure(state="disabled")
            self.aspect_ratio_var.set(False)

        # Ambito di applicazione
        scope_label = ctk.CTkLabel(self, text="Applica a:")
        scope_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.scope_var = ctk.StringVar(value="current")
        
        radio_frame = ctk.CTkFrame(self, fg_color="transparent")
        radio_frame.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkRadioButton(radio_frame, text="Pagina Corrente", variable=self.scope_var, value="current").pack(anchor="w")
        ctk.CTkRadioButton(radio_frame, text="Tutte le Pagine", variable=self.scope_var, value="all").pack(anchor="w")
        ctk.CTkRadioButton(radio_frame, text="Pagine Selezionate", variable=self.scope_var, value="selected").pack(anchor="w")

        # Pulsanti di conferma
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        apply_btn = ctk.CTkButton(button_frame, text="Applica", command=self.apply, fg_color="green")
        apply_btn.pack(side="left", padx=10)
        
        cancel_btn = ctk.CTkButton(button_frame, text="Annulla", command=self.destroy, fg_color="gray")
        cancel_btn.pack(side="left", padx=10)
        
        self.width_var.trace_add("write", self._update_height_from_width)
        self.height_var.trace_add("write", self._update_width_from_height)

    def _update_height_from_width(self, *args):
        if self._is_updating_dims or not self.aspect_ratio_var.get() or self.aspect_ratio is None:
            return
        
        self._is_updating_dims = True
        try:
            width_cm = float(self.width_var.get())
            new_height = width_cm / self.aspect_ratio
            self.height_var.set(f"{new_height:.2f}")
        except (ValueError, ZeroDivisionError):
            pass
        finally:
            self._is_updating_dims = False
            
    def _update_width_from_height(self, *args):
        if self._is_updating_dims or not self.aspect_ratio_var.get() or self.aspect_ratio is None:
            return
            
        self._is_updating_dims = True
        try:
            height_cm = float(self.height_var.get())
            new_width = height_cm * self.aspect_ratio
            self.width_var.set(f"{new_width:.2f}")
        except (ValueError, ZeroDivisionError):
            pass
        finally:
            self._is_updating_dims = False

    def apply(self):
        try:
            self.result = {
                "width_cm": float(self.width_var.get()),
                "height_cm": float(self.height_var.get()),
                "keep_aspect_ratio": self.aspect_ratio_var.get(),
                "scope": self.scope_var.get()
            }
            self.destroy()
        except ValueError:
            messagebox.showerror("Valore non valido", "Inserisci valori numerici per larghezza e altezza.", parent=self)

class DragDropChoiceDialog(ctk.CTkToplevel):
    """
    Dialogo per scegliere se aggiungere o sostituire un file trascinato.
    """
    def __init__(self, master):
        super().__init__(master)
        self.choice = None

        self.title("Azione Drag & Drop")
        self.geometry("350x120")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        label = ctk.CTkLabel(self, text="Un file è già aperto. Cosa vuoi fare?")
        label.pack(pady=10)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=10, fill="x", expand=True)
        button_frame.grid_columnconfigure((0,1,2), weight=1)

        replace_btn = ctk.CTkButton(button_frame, text="Sostituisci File", command=lambda: self.set_choice("replace"))
        replace_btn.grid(row=0, column=0, padx=5)

        add_btn = ctk.CTkButton(button_frame, text="Aggiungi Pagine", command=lambda: self.set_choice("add"))
        add_btn.grid(row=0, column=1, padx=5)
        
        cancel_btn = ctk.CTkButton(button_frame, text="Annulla", command=lambda: self.set_choice("cancel"), fg_color="gray")
        cancel_btn.grid(row=0, column=2, padx=5)

    def set_choice(self, choice):
        self.choice = choice
        self.destroy()

class PDFCheckerApp(ctk.CTkFrame):
    """
    Applicazione per analizzare file PDF, visualizzare, ritagliare,
    eliminare e aggiungere pagine.
    """
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        self.doc = None
        self.doc_path = None
        self.modified_doc = None
        self.active_page_index = -1
        self.ctk_image = None
        self.page_thumbnail_widgets = []
        self.zoom_level = 1.0
        self.pan_start_x = 0
        self.pan_start_y = 0

        # --- Layout a 3 colonne ---
        self.grid_columnconfigure(0, weight=0, minsize=250) 
        self.grid_columnconfigure(1, weight=1, minsize=350)
        self.grid_columnconfigure(2, weight=0, minsize=620) 
        self.grid_rowconfigure(0, weight=1)

        # --- COLONNA 0: Miniature ---
        self.thumbnail_list_frame = ctk.CTkScrollableFrame(self, label_text="Pagine")
        self.thumbnail_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- COLONNA 1: Dettagli e Azioni (Centrale) ---
        center_frame = ctk.CTkFrame(self, fg_color="transparent")
        center_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        center_frame.grid_columnconfigure(0, weight=1)
        center_frame.grid_rowconfigure(0, weight=0) # Info
        center_frame.grid_rowconfigure(1, weight=1) # Azioni + Spazio
        center_frame.grid_rowconfigure(2, weight=0) # Controlli Zoom

        top_info_frame = ctk.CTkFrame(center_frame)
        top_info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        top_info_frame.grid_columnconfigure(1, weight=1)
        
        self.info_vars = {}
        top_prop_labels = ["Nome file:", "Percorso:"]
        for i, label_text in enumerate(top_prop_labels):
            label = ctk.CTkLabel(top_info_frame, text=label_text, anchor="e", font=ctk.CTkFont(weight="bold"))
            label.grid(row=i, column=0, padx=10, pady=2, sticky="ne")
            self.info_vars[label_text] = ctk.StringVar(value="-")
            value_label = ctk.CTkLabel(top_info_frame, textvariable=self.info_vars[label_text], anchor="w", wraplength=400)
            value_label.grid(row=i, column=1, padx=10, pady=2, sticky="w")

        details_actions_frame = ctk.CTkFrame(center_frame, fg_color="transparent")
        details_actions_frame.grid(row=1, column=0, sticky="new")
        
        details_frame = ctk.CTkFrame(details_actions_frame)
        details_frame.pack(fill="x", expand=False, pady=10)
        details_frame.grid_columnconfigure(1, weight=1)
        
        detail_prop_labels = ["MediaBox (cm):", "TrimBox (cm):"]
        for i, label_text in enumerate(detail_prop_labels):
            label = ctk.CTkLabel(details_frame, text=label_text, anchor="e", font=ctk.CTkFont(weight="bold"))
            label.grid(row=i, column=0, padx=10, pady=2, sticky="e")
            self.info_vars[label_text] = ctk.StringVar(value="-")
            value_label = ctk.CTkLabel(details_frame, textvariable=self.info_vars[label_text], anchor="w")
            value_label.grid(row=i, column=1, padx=10, pady=2, sticky="w")

        actions_frame = ctk.CTkFrame(details_actions_frame, fg_color="transparent")
        actions_frame.pack(fill="x", expand=False, pady=10)
        
        self.trim_button = ctk.CTkButton(actions_frame, text="Rifila Pagina Corrente", command=self.crop_to_trimbox, state="disabled")
        self.trim_button.pack(pady=5, fill="x")
        
        self.trim_all_button = ctk.CTkButton(actions_frame, text="Rifila Tutte le Pagine", command=self.crop_all_pages_to_trimbox, state="disabled")
        self.trim_all_button.pack(pady=5, fill="x")

        self.add_pages_button = ctk.CTkButton(actions_frame, text="Aggiungi Pagine da PDF...", command=self.add_pages, state="disabled")
        self.add_pages_button.pack(pady=5, fill="x")
        
        self.extract_pages_button = ctk.CTkButton(actions_frame, text="Estrai Pagine Singole...", command=self.extract_pages, state="disabled")
        self.extract_pages_button.pack(pady=5, fill="x")

        self.add_margins_button = ctk.CTkButton(actions_frame, text="Aggiungi Margini...", command=self.open_add_margins_dialog, state="disabled")
        self.add_margins_button.pack(pady=5, fill="x")
        
        self.resize_button = ctk.CTkButton(actions_frame, text="Ridimensiona Pagine...", command=self.open_resize_dialog, state="disabled")
        self.resize_button.pack(pady=5, fill="x")
        
        self.save_button = ctk.CTkButton(actions_frame, text="Salva PDF Modificato...", command=self.save_modified_pdf, state="disabled", fg_color="green", hover_color="darkgreen")
        self.save_button.pack(pady=(15, 5), fill="x")
        
        zoom_frame = ctk.CTkFrame(center_frame, fg_color="transparent")
        zoom_frame.grid(row=2, column=0, sticky="sew", pady=(20,5))
        zoom_frame.grid_columnconfigure(0, weight=1)

        zoom_out_button = ctk.CTkButton(zoom_frame, text="-", width=40, command=self.zoom_out)
        zoom_out_button.grid(row=0, column=1, padx=5)
        self.zoom_label = ctk.CTkLabel(zoom_frame, text="100%", width=60)
        self.zoom_label.grid(row=0, column=2, padx=5)
        zoom_in_button = ctk.CTkButton(zoom_frame, text="+", width=40, command=self.zoom_in)
        zoom_in_button.grid(row=0, column=3, padx=5)

        # --- COLONNA 2: Anteprima (Destra) ---
        right_frame = ctk.CTkFrame(self, fg_color="transparent")
        right_frame.grid(row=0, column=2, padx=(0, 10), pady=10, sticky="nsew")
        
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        self.preview_container = ctk.CTkFrame(right_frame, width=600, height=600)
        self.preview_container.place(relx=0.5, rely=0.5, anchor="center")
        self.preview_container.pack_propagate(False)

        self.preview_frame = ctk.CTkScrollableFrame(self.preview_container, label_text="")
        self.preview_frame.pack(fill="both", expand=True)
        
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.grid_rowconfigure(0, weight=1)

        self.image_label = ctk.CTkLabel(self.preview_frame, text="", text_color="gray")
        self.image_label.grid(row=0, column=0)
        
        if DND_SUPPORT:
            self.drop_target = self
            self.drop_target.drop_target_register(DND_FILES)
            self.drop_target.dnd_bind('<<Drop>>', self.handle_drop)
            self.image_label.configure(text="Trascina un file PDF qui")
        else:
            self.image_label.configure(text="Funzionalità di trascinamento disabilitata.")

        self.image_label.bind("<ButtonPress-1>", self._on_pan_start)
        self.image_label.bind("<B1-Motion>", self._on_pan_move)
        self.image_label.bind("<ButtonRelease-1>", self._on_pan_end)
        self.image_label.configure(cursor="hand2")

    def _on_pan_start(self, event):
        self.image_label.configure(cursor="fleur")
        self.pan_start_x = event.x
        self.pan_start_y = event.y

    def _on_pan_move(self, event):
        dx = event.x - self.pan_start_x
        dy = event.y - self.pan_start_y
        self.preview_frame._parent_canvas.xview_scroll(-dx, "units")
        self.preview_frame._parent_canvas.yview_scroll(-dy, "units")
        self.pan_start_x = event.x
        self.pan_start_y = event.y

    def _on_pan_end(self, event):
        self.image_label.configure(cursor="hand2")

    def handle_drop(self, event):
        try:
            paths = re.findall(r'\{.*?\}|\S+', event.data)
            cleaned_paths = [p.strip('{}') for p in paths]

            if len(cleaned_paths) > 1:
                messagebox.showwarning("File Multipli", "È possibile trascinare un solo file alla volta.", parent=self)
                return
            
            file_path = cleaned_paths[0]
            
            if not os.path.exists(file_path) or not file_path.lower().endswith('.pdf'):
                messagebox.showwarning("File non valido", f"Il file trascinato non è un PDF valido.", parent=self)
                return

            if self.doc is None:
                self._process_pdf(file_path)
            else:
                dialog = DragDropChoiceDialog(self)
                self.wait_window(dialog)
                choice = dialog.choice
                
                if choice == "replace":
                    self._process_pdf(file_path)
                elif choice == "add":
                    self.add_pages(add_path=file_path)

        except Exception as e:
            messagebox.showerror("Errore Drag & Drop", f"Si è verificato un errore imprevisto.\n\nDettagli: {e}", parent=self)
            self._clear_all()

    def _process_pdf(self, file_path):
        try:
            if self.doc: self.doc.close()
            if self.modified_doc: self.modified_doc.close()

            self.doc = fitz.open(file_path)
            self.doc_path = file_path
            self.modified_doc = None
            self.save_button.configure(state="disabled")
            self.info_vars["Nome file:"].set(os.path.basename(file_path))
            self.info_vars["Percorso:"].set(os.path.dirname(file_path))
            self._populate_thumbnail_list()
        except Exception as e:
            messagebox.showerror("Errore Apertura PDF", f"Impossibile aprire il file PDF.\n\nDettagli: {e}\n\nTraceback:\n{traceback.format_exc()}", parent=self)
            self._clear_all()

    def _populate_thumbnail_list(self):
        for widget in self.thumbnail_list_frame.winfo_children(): widget.destroy()
        self.page_thumbnail_widgets.clear()
        
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return

        num_pages = len(current_doc)
        for page_num in range(num_pages):
            container_frame = ctk.CTkFrame(self.thumbnail_list_frame, fg_color="transparent")
            container_frame.pack(padx=5, pady=5, fill="x")
            container_frame.grid_columnconfigure(0, weight=1)
            container_frame.grid_columnconfigure(1, weight=0)

            thumb_img = self._get_page_thumbnail(current_doc[page_num])
            
            btn = ctk.CTkButton(container_frame, image=thumb_img, text=f"Pag. {page_num + 1}", compound="top", font=ctk.CTkFont(size=11), command=lambda idx=page_num: self._display_page_details(idx), cursor="hand2")
            btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

            actions_subframe = ctk.CTkFrame(container_frame, fg_color="transparent")
            actions_subframe.grid(row=0, column=1, sticky="ns")
            
            move_up_btn = ctk.CTkButton(actions_subframe, text="▲", width=28, command=lambda idx=page_num: self._move_page(idx, 'up'))
            move_up_btn.pack(padx=2, pady=(0, 2), fill="x")
            if page_num == 0:
                move_up_btn.configure(state="disabled")

            move_down_btn = ctk.CTkButton(actions_subframe, text="▼", width=28, command=lambda idx=page_num: self._move_page(idx, 'down'))
            move_down_btn.pack(padx=2, pady=2, fill="x")
            if page_num == num_pages - 1:
                move_down_btn.configure(state="disabled")
            
            if num_pages > 1:
                delete_btn = ctk.CTkButton(actions_subframe, text="X", width=28, fg_color="red", hover_color="darkred", command=lambda idx=page_num: self.delete_page(idx))
                delete_btn.pack(padx=2, pady=(10, 0), fill="x")
            
            self.page_thumbnail_widgets.append({"frame": container_frame, "button": btn})

        if num_pages > 0:
            if self.active_page_index >= num_pages or self.active_page_index < 0:
                self.active_page_index = 0
            self._display_page_details(self.active_page_index)
        else:
            self._clear_details()

    def _move_page(self, page_index, direction):
        self._ensure_modifiable_doc()
        num_pages = len(self.modified_doc)
        new_active_index = -1

        if direction == 'up' and page_index > 0:
            self.modified_doc.move_page(page_index, page_index - 1)
            new_active_index = page_index - 1
        elif direction == 'down' and page_index < num_pages - 1:
            self.modified_doc.move_page(page_index, page_index + 2)
            new_active_index = page_index + 1
        
        if new_active_index != -1:
            self.active_page_index = new_active_index
            self.save_button.configure(state="normal")
            self._populate_thumbnail_list()

    def _get_page_thumbnail(self, page, size=(150, 150)):
        pix = page.get_pixmap(dpi=72)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail(size, Image.Resampling.LANCZOS)
        return ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

    def _display_page_details(self, page_index):
        if page_index < 0 or not (self.modified_doc or self.doc):
            self._clear_details()
            return
        try:
            self.active_page_index = page_index
            current_doc = self.modified_doc if self.modified_doc else self.doc
            
            if page_index >= len(current_doc):
                self._clear_details()
                return

            page = current_doc[page_index]
            
            for i, widget_dict in enumerate(self.page_thumbnail_widgets):
                btn = widget_dict["button"]
                btn.configure(fg_color="transparent" if i != page_index else ctk.ThemeManager.theme["CTkButton"]["fg_color"])
            
            media_box = page.rect
            trim_box = page.trimbox

            self.info_vars["MediaBox (cm):"].set(f"{(media_box.width / 72) * 2.54:.2f} x {(media_box.height / 72) * 2.54:.2f} cm")
            self.info_vars["TrimBox (cm):"].set(f"{(trim_box.width / 72) * 2.54:.2f} x {(trim_box.height / 72) * 2.54:.2f} cm" if trim_box and trim_box.is_valid else "- (non definito)")

            self._fit_zoom_to_view(page)
            self._update_preview_image()
            
            self.trim_button.configure(state="normal" if page.trimbox != page.mediabox and page.trimbox.is_valid else "disabled")
            self.trim_all_button.configure(state="normal")
            self.add_pages_button.configure(state="normal")
            self.extract_pages_button.configure(state="normal")
            self.add_margins_button.configure(state="normal")
            self.resize_button.configure(state="normal")

        except Exception as e:
            messagebox.showerror("Errore Visualizzazione", f"Impossibile mostrare dettagli pagina.\n\nDettagli: {e}", parent=self)

    def _fit_zoom_to_view(self, page):
        try:
            preview_width = 600
            preview_height = 600
            
            preview_width -= 20
            preview_height -= 20

            page_width = page.rect.width
            page_height = page.rect.height

            if page.rotation in [90, 270]:
                page_width, page_height = page_height, page_width

            if page_width == 0 or page_height == 0:
                return

            zoom_x = preview_width / page_width
            zoom_y = preview_height / page_height
            
            self.zoom_level = min(zoom_x, zoom_y)

        except Exception as e:
            self.zoom_level = 1.0

    def zoom_in(self):
        self.zoom_level *= 1.25
        self._update_preview_image()

    def zoom_out(self):
        self.zoom_level /= 1.25
        if self.zoom_level < 0.05: self.zoom_level = 0.05
        self._update_preview_image()

    def _update_preview_image(self):
        if self.active_page_index == -1 or not (self.modified_doc or self.doc):
            return

        current_doc = self.modified_doc if self.modified_doc else self.doc
        page = current_doc[self.active_page_index]
        
        try:
            matrix = fitz.Matrix(self.zoom_level, self.zoom_level)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            self.ctk_image = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.image_label.configure(image=self.ctk_image, text="")
            self.zoom_label.configure(text=f"{self.zoom_level*100:.0f}%")

        except Exception as e:
            self.image_label.configure(image=None, text="Errore anteprima")

    def _ensure_modifiable_doc(self):
        if self.modified_doc is None and self.doc_path:
            self.modified_doc = fitz.open(self.doc_path)

    def delete_page(self, page_index):
        try:
            self._ensure_modifiable_doc()
            self.modified_doc.delete_page(page_index)
            self.save_button.configure(state="normal")
            
            if len(self.modified_doc) > 0:
                if self.active_page_index == page_index:
                    self.active_page_index = max(0, page_index - 1)
                elif self.active_page_index > page_index:
                    self.active_page_index -= 1
            else:
                self.active_page_index = -1
                
            self._populate_thumbnail_list()
        except Exception as e:
            messagebox.showerror("Errore Eliminazione", f"Impossibile eliminare la pagina.\n\nDettagli: {e}", parent=self)

    def add_pages(self, add_path=None):
        if not add_path:
            add_path = filedialog.askopenfilename(parent=self, title="Seleziona PDF da aggiungere", filetypes=[("File PDF", "*.pdf")])
        
        if not add_path:
            return
            
        try:
            with fitz.open(add_path) as doc_to_add:
                if len(doc_to_add) == 1:
                    selected_pages = [0]
                else:
                    dialog = PageSelectionDialog(self, doc_to_add, title="Seleziona Pagine da Aggiungere", button_text="Aggiungi")
                    self.wait_window(dialog)
                    selected_pages = dialog.selected_pages
                
                if not selected_pages:
                    return

                self._ensure_modifiable_doc()
                
                temp_doc = fitz.open() 
                for page_num in selected_pages:
                    temp_doc.insert_pdf(doc_to_add, from_page=page_num, to_page=page_num)
                
                self.modified_doc.insert_pdf(temp_doc)
                temp_doc.close()
                
                self.save_button.configure(state="normal")
                self._populate_thumbnail_list()

        except Exception as e:
            messagebox.showerror("Errore Aggiunta Pagine", f"Impossibile aggiungere le pagine.\n\nDettagli: {e}", parent=self)

    def crop_to_trimbox(self):
        if self.active_page_index == -1: return
        try:
            self._ensure_modifiable_doc()
            page = self.modified_doc[self.active_page_index]
            if page.trimbox and page.trimbox.is_valid and page.trimbox != page.mediabox:
                page.set_cropbox(page.trimbox)
                self._populate_thumbnail_list() 
                self.save_button.configure(state="normal")
                messagebox.showinfo("Rifilatura Applicata", "La pagina è stata rifilata.\nSalva il PDF per rendere le modifiche permanenti.", parent=self)
            else:
                messagebox.showinfo("Nessuna Azione", "TrimBox coincide già con MediaBox o non è valido.", parent=self)
        except Exception as e:
            messagebox.showerror("Errore Rifilatura", f"Impossibile applicare TrimBox.\n\nDettagli: {e}", parent=self)

    def crop_all_pages_to_trimbox(self):
        if not (self.modified_doc or self.doc): return
        try:
            self._ensure_modifiable_doc()
            modified_pages_count = 0
            for page in self.modified_doc:
                if page.trimbox and page.trimbox.is_valid and page.trimbox != page.mediabox:
                    page.set_cropbox(page.trimbox)
                    modified_pages_count += 1
            if modified_pages_count > 0:
                self._populate_thumbnail_list()
                self.save_button.configure(state="normal")
                messagebox.showinfo("Rifilatura Completata", f"{modified_pages_count} pagine sono state rifilate.\nSalva il PDF per rendere le modifiche permanenti.", parent=self)
            else:
                messagebox.showinfo("Nessuna Azione", "Nessuna pagina richiedeva la rifilatura.", parent=self)
        except Exception as e:
            messagebox.showerror("Errore Rifilatura Globale", f"Impossibile applicare TrimBox a tutte le pagine.\n\nDettagli: {e}", parent=self)

    def save_modified_pdf(self):
        if not self.modified_doc: return
        initial_dir = os.path.dirname(self.doc_path)
        initial_file = os.path.splitext(os.path.basename(self.doc_path))[0] + "_modificato.pdf"
        save_path = filedialog.asksaveasfilename(parent=self, title="Salva PDF Modificato", initialdir=initial_dir, initialfile=initial_file, defaultextension=".pdf", filetypes=[("File PDF", "*.pdf")])
        if save_path:
            try:
                self.modified_doc.save(save_path, garbage=4, deflate=True)
                messagebox.showinfo("Salvataggio Completato", f"File salvato con successo in:\n{save_path}", parent=self)
            except Exception as e:
                messagebox.showerror("Errore di Salvataggio", f"Impossibile salvare il file.\n\nDettagli: {e}", parent=self)

    def extract_pages(self):
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return

        dialog = PageSelectionDialog(self, current_doc, title="Seleziona Pagine da Estrarre", button_text="Estrai")
        self.wait_window(dialog)
        selected_pages = dialog.selected_pages

        if not selected_pages: return

        save_folder = filedialog.askdirectory(parent=self, title="Seleziona la cartella dove salvare le pagine estratte")
        
        if not save_folder: return

        try:
            base_filename = os.path.splitext(os.path.basename(self.doc_path))[0]
            num_pages_total = len(current_doc)
            
            for i in selected_pages:
                new_doc = fitz.open()
                new_doc.insert_pdf(current_doc, from_page=i, to_page=i)
                
                page_num_str = str(i + 1).zfill(len(str(num_pages_total)))
                output_path = os.path.join(save_folder, f"{base_filename}_pagina_{page_num_str}.pdf")
                
                new_doc.save(output_path, garbage=4, deflate=True)
                new_doc.close()

            messagebox.showinfo("Estrazione Completata", f"{len(selected_pages)} pagine sono state estratte con successo nella cartella:\n{save_folder}", parent=self)

        except Exception as e:
            messagebox.showerror("Errore Estrazione", f"Si è verificato un errore durante l'estrazione delle pagine.\n\nDettagli: {e}", parent=self)

    def open_add_margins_dialog(self):
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return
        
        dialog = AddMarginsDialog(self)
        self.wait_window(dialog)
        
        if dialog.result:
            self._apply_margins(dialog.result)

    def _apply_margins(self, settings):
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return

        scope = settings["scope"]
        page_indices = []

        if scope == "current":
            if self.active_page_index != -1:
                page_indices.append(self.active_page_index)
        elif scope == "all":
            page_indices = list(range(len(current_doc)))
        elif scope == "selected":
            dialog = PageSelectionDialog(self, current_doc, title="Seleziona Pagine per Margini", button_text="Applica")
            self.wait_window(dialog)
            page_indices = dialog.selected_pages
        
        if not page_indices: return

        try:
            self._ensure_modifiable_doc()
            
            margin_top_points = settings["margin_top"] * MM_TO_PT
            margin_bottom_points = settings["margin_bottom"] * MM_TO_PT
            margin_left_points = settings["margin_left"] * MM_TO_PT
            margin_right_points = settings["margin_right"] * MM_TO_PT

            hex_color = settings["color_hex"].lstrip('#')
            rgb_color = tuple(int(hex_color[i:i+2], 16) / 255.0 for i in (0, 2, 4))

            page_indices.sort(reverse=True)

            for i in page_indices:
                page = self.modified_doc[i]
                original_rect = page.rect
                
                new_rect = fitz.Rect(
                    original_rect.x0 - margin_left_points,
                    original_rect.y0 - margin_top_points,
                    original_rect.x1 + margin_right_points,
                    original_rect.y1 + margin_bottom_points
                )
                
                temp_doc = fitz.open()
                new_page = temp_doc.new_page(width=new_rect.width, height=new_rect.height)
                
                new_page.draw_rect(new_page.rect, color=rgb_color, fill=rgb_color)
                
                target_rect = fitz.Rect(margin_left_points, margin_top_points, new_rect.width - margin_right_points, new_rect.height - margin_bottom_points)
                new_page.show_pdf_page(target_rect, self.modified_doc, i)
                
                self.modified_doc.delete_page(i)
                self.modified_doc.insert_pdf(temp_doc, from_page=0, to_page=0, start_at=i)
                temp_doc.close()

            self.save_button.configure(state="normal")
            self._populate_thumbnail_list()
            messagebox.showinfo("Successo", f"Margini aggiunti a {len(page_indices)} pagina/e.", parent=self)

        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aggiungere i margini.\n\nDettagli: {e}", parent=self)
            
    def open_resize_dialog(self):
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return
        
        current_page = None
        if self.active_page_index != -1:
            current_page = current_doc[self.active_page_index]

        dialog = ResizePageDialog(self, current_page=current_page)
        self.wait_window(dialog)
        
        if dialog.result:
            self._apply_resize(dialog.result)

    def _apply_resize(self, settings):
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return

        scope = settings["scope"]
        page_indices = []

        if scope == "current":
            if self.active_page_index != -1:
                page_indices.append(self.active_page_index)
        elif scope == "all":
            page_indices = list(range(len(current_doc)))
        elif scope == "selected":
            dialog = PageSelectionDialog(self, current_doc, title="Seleziona Pagine da Ridimensionare", button_text="Applica")
            self.wait_window(dialog)
            page_indices = dialog.selected_pages
        
        if not page_indices: return

        try:
            self._ensure_modifiable_doc()
            
            new_width_pt = settings["width_cm"] * CM_TO_PT
            new_height_pt = settings["height_cm"] * CM_TO_PT

            page_indices.sort(reverse=True)

            for i in page_indices:
                old_page = self.modified_doc[i]
                
                temp_doc = fitz.open()
                new_page = temp_doc.new_page(width=new_width_pt, height=new_height_pt)
                
                target_rect = new_page.rect
                if settings["keep_aspect_ratio"]:
                    zoom_x = new_width_pt / old_page.rect.width
                    zoom_y = new_height_pt / old_page.rect.height
                    scale = min(zoom_x, zoom_y)
                    
                    final_w = old_page.rect.width * scale
                    final_h = old_page.rect.height * scale
                    
                    pos_x = (new_width_pt - final_w) / 2
                    pos_y = (new_height_pt - final_h) / 2
                    target_rect = fitz.Rect(pos_x, pos_y, pos_x + final_w, pos_y + final_h)

                new_page.show_pdf_page(target_rect, self.modified_doc, i)
                
                self.modified_doc.delete_page(i)
                self.modified_doc.insert_pdf(temp_doc, from_page=0, to_page=0, start_at=i)
                temp_doc.close()
                
            self.save_button.configure(state="normal")
            self._populate_thumbnail_list()
            messagebox.showinfo("Successo", f"Ridimensionamento applicato a {len(page_indices)} pagina/e.", parent=self)

        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile ridimensionare le pagine.\n\nDettagli: {e}", parent=self)


    def _clear_all(self):
        if self.doc: self.doc.close()
        if self.modified_doc: self.modified_doc.close()
        self.doc = None; self.doc_path = None; self.modified_doc = None
        self.active_page_index = -1
        for widget in self.thumbnail_list_frame.winfo_children(): widget.destroy()
        self.page_thumbnail_widgets.clear()
        self._clear_details()

    def _clear_details(self):
        for key in ["Nome file:", "Percorso:", "MediaBox (cm):", "TrimBox (cm):"]:
            if key in self.info_vars: self.info_vars[key].set("-")
        self.image_label.configure(image=None, text="Trascina un file PDF qui")
        self.trim_button.configure(state="disabled")
        self.trim_all_button.configure(state="disabled")
        self.add_pages_button.configure(state="disabled")
        self.extract_pages_button.configure(state="disabled")
        self.add_margins_button.configure(state="disabled")
        self.resize_button.configure(state="disabled")
        self.save_button.configure(state="disabled")

def create_tab(tab_view):
    tab_name = "Controllo PDF"
    tab = tab_view.add(tab_name)
    app_instance = PDFCheckerApp(master=tab)
    return tab_name, app_instance

