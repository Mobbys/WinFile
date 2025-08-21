# app_liste_anteprime.py - v2.5
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk, Menu
import os
import threading
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import io
from collections import defaultdict
import traceback
import webbrowser
import tempfile
import base64
import csv
import html

# Import per la gestione avanzata degli appunti su Windows
import ctypes
from ctypes import wintypes

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4, landscape

# Disabilita il limite di dimensione per le immagini grandi
Image.MAX_IMAGE_PIXELS = None

# --- COSTANTI ---
SUPPORTED_EXTENSIONS = ('.jpg', '.jpeg', '.tif', '.tiff', '.pdf', '.ai')
DEFAULT_DPI = 96
PREVIEW_SIZE = (300, 300)


class ExportOptionsWindow(ctk.CTkToplevel):
    """
    Finestra di dialogo (convertita in CustomTkinter) per le opzioni di esportazione.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Opzioni di Esportazione")
        # Modifica: Aumentata la larghezza per contenere tutto il testo
        self.geometry("350x220")
        self.transient(parent)
        self.grab_set()

        self.result = None

        self.orientation = ctk.StringVar(value="portrait")
        self.columns = ctk.IntVar(value=4)

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        orientation_frame = ctk.CTkFrame(main_frame)
        orientation_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(orientation_frame, text="Orientamento Pagina:").pack(side="left")
        ctk.CTkRadioButton(orientation_frame, text="Orizzontale", variable=self.orientation, value="landscape").pack(side="left", padx=5, expand=True)
        ctk.CTkRadioButton(orientation_frame, text="Verticale", variable=self.orientation, value="portrait").pack(side="left", padx=5, expand=True)

        columns_frame = ctk.CTkFrame(main_frame)
        columns_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(columns_frame, text="Colonne per Riga:").pack(side="left", padx=5)
        ctk.CTkEntry(columns_frame, textvariable=self.columns, width=60).pack(side="left")

        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(pady=10, fill="x")
        ctk.CTkButton(button_frame, text="OK", command=self.on_ok).pack(side="left", padx=5, expand=True)
        ctk.CTkButton(button_frame, text="Annulla", command=self.on_cancel, fg_color="gray").pack(side="left", padx=5, expand=True)

    def on_ok(self):
        try:
            cols = self.columns.get()
            if not (1 <= cols <= 20):
                raise ValueError
            self.result = {"orientation": self.orientation.get(), "columns": cols}
            self.destroy()
        except ValueError:
            messagebox.showerror("Input non valido", "Il numero di colonne deve essere un intero tra 1 e 20.", parent=self)

    def on_cancel(self):
        self.result = None
        self.destroy()

class FileScannerApp(ctk.CTkFrame):
    """
    Classe principale dell'applicazione. Il Drag & Drop è gestito globalmente da winfile.py.
    """
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        self.status_text = ctk.StringVar(value="Trascina file o cartelle qui, oppure usa il pulsante.")
        self.preview_image = None
        self.scan_results = []
        self.is_scanning = False
        
        self.bottom_button_groups = [] # Lista per i gruppi di pulsanti in basso

        self.create_widgets()
        self.create_context_menu()
        self.style_treeview()
        
        # Associa l'evento di ridimensionamento al frame dei controlli in basso
        self.bottom_controls_frame.bind('<Configure>', self._rearrange_button_groups)


    def handle_drop(self, event):
        """
        Questo metodo viene chiamato da winfile.py quando dei file vengono
        rilasciati sulla finestra dell'applicazione.
        """
        if self.is_scanning:
            return
        
        path_string = event.data
        
        try:
            paths = self.tk.splitlist(path_string)
        except Exception:
            paths = []

        if paths:
            self.run_scan(paths)
        else:
            self.status_text.set("Nessun file o cartella valida trascinata.")
            print(f"DEBUG: Impossibile interpretare i percorsi da event.data: '{event.data}'")

    def _rearrange_button_groups(self, event=None):
        """
        Riorganizza i gruppi di pulsanti in basso in modo reattivo (a capo)
        in base alla larghezza disponibile del frame contenitore.
        """
        if not self.winfo_viewable():
            return
            
        available_width = self.bottom_controls_frame.winfo_width()
        padding_x = 2
        padding_y = 2
        
        cursor_x = padding_x
        cursor_y = padding_y
        row_height = 0

        for group in self.bottom_button_groups:
            group_width = group.winfo_reqwidth()
            group_height = group.winfo_reqheight()

            if cursor_x + group_width + padding_x > available_width and cursor_x > padding_x:
                # Sposta il gruppo a una nuova riga
                cursor_y += row_height + padding_y
                cursor_x = padding_x
                row_height = 0

            group.place(x=cursor_x, y=cursor_y)
            cursor_x += group_width + padding_x
            
            if group_height > row_height:
                row_height = group_height

        # Adatta l'altezza del frame contenitore
        required_height = cursor_y + row_height + padding_y
        if self.bottom_controls_frame.winfo_reqheight() != required_height:
             self.bottom_controls_frame.configure(height=required_height)


    def create_widgets(self):
        # --- Frame Superiore (solo gestione lista) ---
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        self.select_button = ctk.CTkButton(top_frame, text="Aggiungi Cartella", command=self.select_folder_dialog)
        self.select_button.pack(side="left", padx=5, pady=5)
        self.clear_button = ctk.CTkButton(top_frame, text="Svuota Lista", command=self.clear_results, state="disabled")
        self.clear_button.pack(side="left", padx=5, pady=5)

        # --- Frame Centrale (Lista e Anteprima) ---
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        main_frame.grid_columnconfigure(0, weight=3); main_frame.grid_columnconfigure(1, weight=1); main_frame.grid_rowconfigure(0, weight=1)

        tree_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("filename", "dimensions_cm", "path")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings")
        self.tree.column("#0", width=30, stretch=False, anchor="center")
        self.tree.heading("#0", text="")
        self.tree.heading("filename", text="Nome File / Pagina"); self.tree.heading("dimensions_cm", text="Dimensioni (cm)"); self.tree.heading("path", text="Percorso")
        self.tree.column("filename", width=300); self.tree.column("dimensions_cm", width=150, anchor="center"); self.tree.column("path", width=350)
        scrollbar = ctk.CTkScrollbar(tree_frame, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_item_select)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", self.open_selected_file) # Aggiunto binding per doppio click
        
        preview_frame = ctk.CTkFrame(main_frame)
        preview_frame.grid(row=0, column=1, sticky="nsew")
        self.preview_label = ctk.CTkLabel(preview_frame, text="")
        self.preview_label.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Frame Inferiore (Controlli Copia ed Esporta) ---
        self.bottom_controls_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_controls_frame.pack(fill="x", padx=10, pady=5)

        # Gruppo Copia
        copy_frame = ctk.CTkFrame(self.bottom_controls_frame)
        ctk.CTkLabel(copy_frame, text="Copia").pack(anchor="w", pady=(2, 2), padx=10)
        self.copy_all_button = ctk.CTkButton(copy_frame, text="Copia Tabella", command=self.copy_all_to_clipboard, state="disabled")
        self.copy_all_button.pack(side="left", padx=5, pady=(0, 5))
        self.copy_formatted_button = ctk.CTkButton(copy_frame, text="Copia Formattata", command=lambda: self.copy_formatted_to_clipboard(selection_mode=False), state="disabled")
        self.copy_formatted_button.pack(side="left", padx=5, pady=(0, 5))
        self.bottom_button_groups.append(copy_frame)

        # Gruppo Esporta/Stampa
        export_frame = ctk.CTkFrame(self.bottom_controls_frame)
        ctk.CTkLabel(export_frame, text="Esporta / Stampa").pack(anchor="w", pady=(2, 2), padx=10)
        self.print_button = ctk.CTkButton(export_frame, text="Stampa Tabella", command=lambda: self.print_table(selection_mode=False), state="disabled")
        self.print_button.pack(side="left", padx=5, pady=(0, 5))
        self.export_html_button = ctk.CTkButton(export_frame, text="Anteprima HTML", command=self.export_to_html, state="disabled")
        self.export_html_button.pack(side="left", padx=5, pady=(0, 5))
        self.export_csv_button = ctk.CTkButton(export_frame, text="Esporta in CSV", command=lambda: self.export_to_csv(selection_mode=False), state="disabled")
        self.export_csv_button.pack(side="left", padx=5, pady=(0, 5))
        self.export_pdf_button = ctk.CTkButton(export_frame, text="Esporta in PDF", command=self.export_to_pdf, state="disabled")
        self.export_pdf_button.pack(side="left", padx=5, pady=(0, 5))
        self.bottom_button_groups.append(export_frame)

        # --- Barra di Stato ---
        status_bar = ctk.CTkFrame(self, height=30)
        status_bar.pack(side="bottom", fill="x", padx=10, pady=(5, 10))
        status_label = ctk.CTkLabel(status_bar, textvariable=self.status_text, anchor="w")
        status_label.pack(side="left", padx=10)

    def _lock_ui(self):
        """Disabilita i pulsanti durante un'operazione lunga per evitare input multipli."""
        self.select_button.configure(state="disabled")
        self.clear_button.configure(state="disabled")
        self.copy_all_button.configure(state="disabled")
        self.copy_formatted_button.configure(state="disabled")
        self.print_button.configure(state="disabled")
        self.export_html_button.configure(state="disabled")
        self.export_pdf_button.configure(state="disabled")
        self.export_csv_button.configure(state="disabled")

    def _unlock_ui(self):
        """Riabilita i pulsanti dopo un'operazione, rispettando lo stato attuale (es. se ci sono risultati)."""
        self.select_button.configure(state="normal")
        state = "normal" if self.scan_results else "disabled"
        self.clear_button.configure(state=state)
        self.copy_all_button.configure(state=state)
        self.copy_formatted_button.configure(state=state)
        self.print_button.configure(state=state)
        self.export_html_button.configure(state=state)
        self.export_pdf_button.configure(state=state)
        self.export_csv_button.configure(state=state)

    def create_context_menu(self):
        font_size = 13
        self.option_add("*Menu.font", ("Segoe UI", font_size))
        
        self.context_menu = Menu(self, tearoff=0)
        self.context_menu.add_command(label="Copia selezione negli appunti", command=self.copy_selection_to_clipboard)
        self.context_menu.add_command(label="Copia selezione formattata", command=lambda: self.copy_formatted_to_clipboard(selection_mode=True))
        self.context_menu.add_command(label="Stampa selezione", command=lambda: self.print_table(selection_mode=True))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Esporta selezione in CSV", command=lambda: self.export_to_csv(selection_mode=True))
        self.context_menu.add_command(label="Crea Anteprima HTML...", command=lambda: self.export_to_html(selection_mode=True))
        self.context_menu.add_command(label="Esporta selezione in PDF", command=lambda: self.export_to_pdf(selection_mode=True))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Rimuovi selezionati", command=self.remove_selected_items)

    def show_context_menu(self, event):
        if self.tree.selection():
            self.context_menu.post(event.x_root, event.y_root)

    def style_treeview(self):
        style = ttk.Style()
        
        font_size = 13
        row_height = 28
        
        if ctk.get_appearance_mode() == "Dark":
            bg_color, text_color, field_bg_color = "#2B2B2B", "white", "#343638"
            selected_color, odd_row, even_row = "#1f6aa5", "#242424", "#2B2B2B"
        else:
            bg_color, text_color, field_bg_color = "white", "black", "#EAEAEA"
            selected_color, odd_row, even_row = "#3484d0", "#F7F7F7", "white"
        
        style.theme_use("default")
        
        style.configure("Treeview", 
                        background=bg_color, 
                        foreground=text_color, 
                        fieldbackground=field_bg_color, 
                        borderwidth=0,
                        font=("Segoe UI", font_size),
                        rowheight=row_height)
                        
        style.map('Treeview', background=[('selected', selected_color)])
        
        style.configure("Treeview.Heading", 
                        background=field_bg_color, 
                        foreground=text_color, 
                        relief="flat",
                        font=("Segoe UI", font_size + 1, "bold"))
                        
        style.map("Treeview.Heading", background=[('active', selected_color)])
        
        self.tree.tag_configure('oddrow', background=odd_row)
        self.tree.tag_configure('evenrow', background=even_row)

    def get_file_details(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            details = {"filename": os.path.basename(file_path), "type": ext.replace('.', '').upper(), "path": os.path.dirname(file_path)}
            
            if ext in ('.jpg', '.jpeg', '.tif', '.tiff'):
                with Image.open(file_path) as img:
                    width_px, height_px = img.size; dpi_x, dpi_y = img.info.get('dpi', (DEFAULT_DPI, DEFAULT_DPI))
                    if dpi_x == 0: dpi_x = DEFAULT_DPI
                    if dpi_y == 0: dpi_y = DEFAULT_DPI
                    width_cm = (width_px/dpi_x)*2.54; height_cm = (height_px/dpi_y)*2.54
                    details.update({"page_count": 1, "pages_details": [{"dimensions_cm": f"{width_cm:.2f} x {height_cm:.2f}", "width_cm": width_cm, "height_cm": height_cm}], "dpi_str": f"{int(dpi_x)} DPI"})
                    return details
            
            elif ext in ('.pdf', '.ai'):
                try:
                    with fitz.open(file_path) as doc:
                        if len(doc) == 0: raise ValueError("Documento vuoto")
                        pages_details = []
                        for page in doc:
                            rect = page.rect; width_cm = (rect.width/72)*2.54; height_cm = (rect.height/72)*2.54
                            pages_details.append({"dimensions_cm": f"{width_cm:.2f} x {height_cm:.2f}", "width_cm": width_cm, "height_cm": height_cm})
                        details.update({"type": "AI" if ext == '.ai' else "PDF", "page_count": doc.page_count, "pages_details": pages_details})
                        return details
                except Exception:
                    details.update({"type": "AI (Non compatibile)" if ext == '.ai' else "PDF (Danneggiato)", "page_count": 1, "pages_details": [{"dimensions_cm": "Non rilevabili", "width_cm": 0, "height_cm": 0}]})
                    return details
                    
        except Exception as e:
            print(f"Errore generale nell'analisi del file {file_path}: {e}")
            return {"filename": os.path.basename(file_path), "type": "ERRORE", "path": os.path.dirname(file_path), "page_count": 1, "pages_details": [{"dimensions_cm": "Errore lettura", "width_cm": 0, "height_cm": 0}]}

    def process_paths(self, paths):
        self.status_text.set("Scansione in corso...")
        self._lock_ui()
        found_files = []
        for path in paths:
            scan_root = os.path.normpath(path)
            if os.path.isdir(scan_root):
                for root_dir, _, files in os.walk(scan_root):
                    for file in files:
                        if file.lower().endswith(SUPPORTED_EXTENSIONS):
                            full_path = os.path.join(root_dir, file)
                            details = self.get_file_details(full_path)
                            if details:
                                details['scan_root'] = scan_root
                                found_files.append(details)
            elif os.path.isfile(scan_root) and scan_root.lower().endswith(SUPPORTED_EXTENSIONS):
                details = self.get_file_details(scan_root)
                if details:
                    details['scan_root'] = os.path.dirname(scan_root)
                    found_files.append(details)
        self.after(0, self.add_scan_results, found_files)

    def add_scan_results(self, new_files):
        self.scan_results.extend(new_files)
        
        unique_results = []; seen_paths = set()
        for item in self.scan_results:
            full_path = os.path.join(item['path'], item['filename'])
            if full_path not in seen_paths:
                unique_results.append(item); seen_paths.add(full_path)
        self.scan_results = unique_results
        
        self.repopulate_treeview()

    def _get_display_path(self, file_info):
        normalized_path = os.path.normpath(file_info['path'])
        scan_root = os.path.normpath(file_info.get('scan_root', ''))
        
        if not scan_root: return os.path.basename(normalized_path)
        try:
            if os.path.splitdrive(normalized_path)[0].upper() == os.path.splitdrive(scan_root)[0].upper():
                relative_path = os.path.relpath(normalized_path, scan_root)
                base_name = os.path.basename(scan_root)
                return os.path.join(base_name, relative_path) if relative_path != '.' else base_name
            else: return os.path.basename(normalized_path)
        except ValueError: return os.path.basename(normalized_path)

    def repopulate_treeview(self):
        """
        Popola il Treeview con una struttura gerarchica per i file multi-pagina.
        """
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        total_rows = 0
        for file_index, file_info in enumerate(self.scan_results):
            page_count = file_info.get('page_count', 1)
            tag = 'evenrow' if file_index % 2 == 0 else 'oddrow'
            
            if page_count > 1:
                parent_id = f"file_{file_index}"
                filename_display = f"{file_info['filename']} ({page_count} pagine)"
                display_path = self._get_display_path(file_info).replace('\\', '/')
                self.tree.insert("", "end", iid=parent_id, values=(filename_display, "Multi-pagina", display_path), tags=(tag,), open=False)
                
                for page_num in range(page_count):
                    page_details = file_info["pages_details"][page_num]
                    child_filename_display = f"   Pagina {page_num + 1}"
                    child_dimensions_display = page_details["dimensions_cm"]
                    child_id = f"{file_index}_{page_num}"
                    self.tree.insert(parent_id, "end", iid=child_id, values=(child_filename_display, child_dimensions_display, ""), tags=(tag,))
                total_rows += page_count
            else:
                page_details = file_info["pages_details"][0]
                filename_display = file_info['filename']
                display_path = self._get_display_path(file_info).replace('\\', '/')
                dimensions_display = page_details["dimensions_cm"]
                item_id = f"{file_index}_0"
                self.tree.insert("", "end", iid=item_id, values=(filename_display, dimensions_display, display_path), tags=(tag,))
                total_rows += 1

        self.status_text.set(f"Scansione completata. Trovati {len(self.scan_results)} file ({total_rows} pagine/elementi).")
        self.preview_label.configure(image=None)
        
        self._unlock_ui()
        self.is_scanning = False

    def select_folder_dialog(self):
        if self.is_scanning: return
        folder = filedialog.askdirectory(parent=self)
        if folder: self.run_scan([folder])

    def run_scan(self, paths):
        self.is_scanning = True
        self.status_text.set("Avvio scansione...")
        scan_thread = threading.Thread(target=self.process_paths, args=(paths,), daemon=True)
        scan_thread.start()

    def clear_results(self):
        self.scan_results = []
        for i in self.tree.get_children(): self.tree.delete(i)
        self.status_text.set("Lista svuotata. Pronto per una nuova scansione.")
        self.preview_label.configure(image=None)
        self._unlock_ui()

    def on_item_select(self, event):
        """
        Gestisce la selezione di un elemento nella lista per mostrare l'anteprima.
        Funziona sia per le righe madri che per le pagine figlie.
        """
        selected_items = self.tree.selection()
        if not selected_items: return
        
        selected_id = selected_items[0]
        
        if selected_id.startswith("file_"):
            file_index_str = selected_id.split('_')[1]
            item_index = int(file_index_str)
            page_to_show = 0
        else:
            file_index_str, page_num_str = selected_id.split('_')
            item_index = int(file_index_str)
            page_to_show = int(page_num_str)
        
        selected_data = self.scan_results[item_index]
        full_path = os.path.join(selected_data['path'], selected_data['filename'])
        
        try:
            if selected_data['type'] in ('PDF', 'AI'):
                with fitz.open(full_path) as doc:
                    page = doc.load_page(page_to_show); pix = page.get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            else:
                img = Image.open(full_path)
            
            img.thumbnail(PREVIEW_SIZE, Image.Resampling.LANCZOS)
            self.preview_image = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
            self.preview_label.configure(image=self.preview_image)

        except Exception as e:
            self.status_text.set(f"Impossibile generare anteprima per {selected_data['filename']}.")
            print(f"Errore anteprima: {e}"); self.preview_label.configure(image=None)

    def open_selected_file(self, event):
        """Apre il file selezionato con l'applicazione predefinita del sistema."""
        selected_id = self.tree.focus()
        if not selected_id:
            return

        if selected_id.startswith("file_"):
            file_index = int(selected_id.split('_')[1])
        else:
            file_index = int(selected_id.split('_')[0])

        try:
            selected_data = self.scan_results[file_index]
            full_path = os.path.join(selected_data['path'], selected_data['filename'])
            
            if os.path.exists(full_path):
                os.startfile(os.path.realpath(full_path))
            else:
                messagebox.showerror("Errore", f"File non trovato:\n{full_path}", parent=self)
        except IndexError:
            print(f"Indice non valido per la selezione: {selected_id}")
        except Exception as e:
            messagebox.showerror("Errore Apertura File", f"Impossibile aprire il file.\n\nDettagli: {e}", parent=self)

    def remove_selected_items(self):
        """Rimuove i file selezionati dalla lista."""
        selected_iids = self.tree.selection()
        if not selected_iids:
            messagebox.showinfo("Informazione", "Nessun file selezionato da rimuovere.", parent=self)
            return

        indices_to_remove = set()
        for iid in selected_iids:
            if iid.startswith("file_"):
                index = int(iid.split('_')[1])
            else:
                index = int(iid.split('_')[0])
            indices_to_remove.add(index)
        
        # Ricostruisce la lista dei risultati escludendo gli indici da rimuovere
        self.scan_results = [item for i, item in enumerate(self.scan_results) if i not in indices_to_remove]
        
        # Ripopola la vista ad albero
        self.repopulate_treeview()


    def get_pages_for_selection(self, selection_mode=False):
        """
        Restituisce una lista di dizionari, ognuno rappresentante una pagina
        da esportare. Contiene i dati del file originale e il numero di pagina.
        """
        pages_to_export = []
        
        if not selection_mode:
            for file_index, file_info in enumerate(self.scan_results):
                for page_num in range(file_info.get('page_count', 1)):
                    pages_to_export.append({'file_info': file_info, 'page_num': page_num})
            return pages_to_export

        selected_iids = self.tree.selection()
        if not selected_iids:
            return []

        unique_pages = set()

        for iid in selected_iids:
            if iid.startswith("file_"):
                file_index = int(iid.split('_')[1])
                file_info = self.scan_results[file_index]
                for page_num in range(file_info.get('page_count', 1)):
                    unique_pages.add((file_index, page_num))
            else:
                file_index, page_num = map(int, iid.split('_'))
                unique_pages.add((file_index, page_num))
                
        for file_index, page_num in sorted(list(unique_pages)):
            pages_to_export.append({'file_info': self.scan_results[file_index], 'page_num': page_num})
            
        return pages_to_export

    def _get_all_data_rows(self):
        """Funzione di supporto per ottenere tutti gli ID delle righe di dati (pagine o file singoli)."""
        data_rows = []
        for parent_id in self.tree.get_children(""):
            if parent_id.startswith("file_"):
                data_rows.extend(self.tree.get_children(parent_id))
            else:
                data_rows.append(parent_id)
        return data_rows

    def _get_data_rows_from_selection(self, selected_iids):
        """Funzione di supporto per ottenere gli ID delle righe di dati da una selezione."""
        data_rows = []
        for iid in selected_iids:
            if iid.startswith("file_"):
                data_rows.extend(self.tree.get_children(iid))
            else:
                data_rows.append(iid)
        return sorted(list(set(data_rows)))

    def copy_all_to_clipboard(self):
        if not self.scan_results: return
        header = ["Nome File / Pagina", "Dimensioni (cm)", "Percorso"]
        lines = ["\t".join(header)]
        
        all_data_rows = self._get_all_data_rows()
        
        for item_id in all_data_rows:
            values = list(self.tree.item(item_id)['values'])
            if not values[2]:
                file_index = int(item_id.split('_')[0])
                file_info = self.scan_results[file_index]
                values[2] = self._get_display_path(file_info).replace('\\', '/')
            lines.append("\t".join(map(str, values)))
        
        self.clipboard_clear(); self.clipboard_append("\n".join(lines))
        self.status_text.set("Tabella copiata negli appunti.")

    def copy_selection_to_clipboard(self):
        selected_items = self.tree.selection()
        if not selected_items: return
        header = ["Nome File / Pagina", "Dimensioni (cm)", "Percorso"]
        lines = ["\t".join(header)]
        
        items_to_copy = self._get_data_rows_from_selection(selected_items)
        
        for item_id in items_to_copy:
            values = list(self.tree.item(item_id)['values'])
            if not values[2]:
                file_index = int(item_id.split('_')[0])
                file_info = self.scan_results[file_index]
                values[2] = self._get_display_path(file_info).replace('\\', '/')
            lines.append("\t".join(map(str, values)))

        self.clipboard_clear(); self.clipboard_append("\n".join(lines))
        self.status_text.set(f"Selezione copiata negli appunti ({len(lines)-1} righe).")

    def print_table(self, selection_mode=False):
        pages_to_print = self.get_pages_for_selection(selection_mode)
        if not pages_to_print:
            messagebox.showinfo("Informazione", "Nessuna riga da stampare.", parent=self); return
        
        html_string = self._generate_html_table_with_totals(pages_to_print)

        html_content = f"""
        <!DOCTYPE html><html><head><meta charset='UTF-8'><title>Stampa Tabella</title>
        <style>
            @media print {{ @page {{ margin: 1.5cm; }} body {{ font-family: sans-serif; -webkit-print-color-adjust: exact; }}
            table {{ width: 100%; border-collapse: collapse; }} th, td {{ border: 1px solid black; padding: 5px; text-align: left; }}
            th {{ background-color: #e0e0e0 !important; }} tfoot {{ background-color: #f0f0f0 !important; font-weight: bold;}}
        </style></head><body onload="window.print()">{html_string}</body></html>"""
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.html', encoding='utf-8') as f:
                f.write(html_content)
            webbrowser.open(f'file://{os.path.realpath(f.name)}')
            self.status_text.set("Apertura anteprima di stampa nel browser...")
        except Exception as e:
            traceback.print_exc()
            self.status_text.set("Errore durante la preparazione della stampa.")
            messagebox.showerror("Errore di Stampa", f"Impossibile aprire l'anteprima di stampa.\nErrore: {e}", parent=self)

    def export_to_csv(self, selection_mode=False):
        pages_to_export = self.get_pages_for_selection(selection_mode)
        if not pages_to_export:
            messagebox.showinfo("Informazione", "Nessun dato da esportare.", parent=self)
            return
        
        self._lock_ui()
        file_path = filedialog.asksaveasfilename(parent=self, defaultextension=".csv", filetypes=[("File CSV", "*.csv")], title="Salva lista come CSV")
        
        if not file_path:
            self.status_text.set("Esportazione CSV annullata.")
            self._unlock_ui()
            return

        self.status_text.set("Creazione del file CSV in corso...")
        self.update_idletasks()
        
        thread = threading.Thread(target=self._build_csv_thread, args=(file_path, pages_to_export), daemon=True)
        thread.start()

    def _build_csv_thread(self, file_path, pages_to_export):
        try:
            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(["Nome File", "Tipo", "Pagina", "Dimensioni (cm)", "Area (m²)", "Percorso Relativo"])
                
                for page_data in pages_to_export:
                    item = page_data['file_info']
                    page_num = page_data['page_num']
                    
                    display_path = self._get_display_path(item).replace('\\', '/')
                    page_details = item['pages_details'][page_num]
                    page_count = item.get('page_count', 1)
                    area_sqm = (page_details.get('width_cm', 0) * page_details.get('height_cm', 0)) / 10000
                    page_str = f"{page_num + 1} di {page_count}"
                    writer.writerow([item['filename'], item['type'], page_str, page_details['dimensions_cm'], f"{area_sqm:.4f}", display_path])
            
            self.after(0, self.on_csv_success, file_path)
        except Exception as e:
            self.after(0, self.on_csv_error, e)

    def on_csv_success(self, file_path):
        self.status_text.set("File CSV creato con successo.")
        self._unlock_ui()
        self.update()
        messagebox.showinfo("Successo", f"Lista esportata con successo in:\n{file_path}", parent=self)

    def on_csv_error(self, e):
        traceback.print_exc()
        self.status_text.set("Errore durante la creazione del CSV.")
        self._unlock_ui()
        self.update()
        messagebox.showerror("Errore", f"Impossibile creare il file CSV.\nErrore: {e}", parent=self)

    def export_to_html(self, selection_mode=False):
        pages_to_export = self.get_pages_for_selection(selection_mode)
        if not pages_to_export:
            messagebox.showinfo("Informazione", "Nessun dato da esportare.", parent=self)
            return

        self._lock_ui()
        self.status_text.set("Preparazione anteprima HTML in corso...")
        self.update_idletasks()
        
        thread = threading.Thread(target=self._build_html_thread, args=(pages_to_export,), daemon=True)
        thread.start()

    def _build_html_thread(self, pages_to_export):
        try:
            html_content = self._generate_html_content(pages_to_export)
            
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.html', encoding='utf-8') as f:
                temp_file_path = f.name
                f.write(html_content)
            
            self.after(0, self.on_html_success, temp_file_path)
        except Exception as e:
            self.after(0, self.on_html_error, e)

    def on_html_success(self, file_path):
        try:
            # Usa os.startfile su Windows per un'apertura più affidabile
            if os.name == 'nt':
                os.startfile(os.path.realpath(file_path))
            else:
                # Fallback per altri sistemi operativi (macOS, Linux)
                webbrowser.open(f'file://{os.path.realpath(file_path)}')
            self.status_text.set("Anteprima HTML creata con successo.")
        except Exception as e:
            self.status_text.set("Errore nell'apertura dell'anteprima.")
            messagebox.showerror("Errore Apertura", f"Impossibile aprire il file HTML nel browser.\n\nDettagli: {e}", parent=self)
        finally:
            self._unlock_ui()
            self.update()

    def on_html_error(self, e):
        traceback.print_exc()
        self.status_text.set("Errore durante la creazione dell'anteprima HTML.")
        self._unlock_ui()
        self.update()
        messagebox.showerror("Errore", f"Impossibile creare l'anteprima.\nErrore: {e}", parent=self)

    def _generate_html_content(self, pages_to_export):
        grouped_pages = defaultdict(list)
        for page in pages_to_export:
            grouped_pages[page['file_info']['scan_root']].append(page)
            
        file_colors = ['#DB4437', '#4285F4', '#F4B400', '#0F9D58', '#AB47BC']
        
        css = """<style id="page-orientation-style"></style><style>@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');:root{--list-thumbnail-size:80px;--grid-item-width:200px}body{font-family:'Roboto',sans-serif;margin:0;background-color:#f4f4f4}.controls{display:flex;flex-wrap:wrap;align-items:center;gap:15px;margin-bottom:20px;background:#fff;padding:10px 15px;border-radius:8px;box-shadow:0 2px 4px #0000001a;position:sticky;top:10px;z-index:1000}#printable-content{margin:20px auto;padding:15mm;background:#fff;box-shadow:0 0 10px #0000001a;box-sizing:border-box}.controls button,.controls select{padding:8px 12px;font-size:14px;background-color:#e9e9e9;color:#333;border:1px solid #ccc;border-radius:5px;cursor:pointer;font-weight:700}#search-box{padding:8px;border:1px solid #ccc;border-radius:5px;width:200px}.print-button{background-color:#4285f4;color:#fff;border-color:#4285f4}.control-group{display:flex;align-items:center;gap:5px}.view-switcher button.active{background-color:#4285f4;color:#fff;border-color:#4285f4}.control-group button.size-btn{width:35px;height:35px;font-size:18px;line-height:1}.folder-container{margin-bottom:20px}.folder-header{background-color:#e0e0e0;padding:10px;border-left:5px solid #4285f4;overflow:hidden}.folder-path{font-weight:700;font-size:1.2em;float:left}.folder-stats{float:right;font-size:.9em;color:#555;line-height:1.5em}.grid-container{display:flex;flex-wrap:wrap;justify-content:center;gap:15px;margin-top:15px}.grid-container .item{width:var(--grid-item-width);display:flex;flex-direction:column;border:1px solid #ddd;border-radius:5px;padding:10px;text-align:center;background-color:#fff;page-break-inside:avoid;position:relative;overflow:hidden}.grid-container .item-img{max-width:100%;height:auto;border-radius:3px;object-fit:contain}.grid-container .item-info{text-align:center;flex-grow:1}.list-container{display:flex;flex-direction:column;gap:8px;margin-top:15px}.list-container .item{display:flex;align-items:center;border:1px solid #ddd;border-radius:5px;padding:8px;background-color:#fff;page-break-inside:avoid;position:relative}.list-container .item-img{width:var(--list-thumbnail-size);height:var(--list-thumbnail-size);object-fit:contain;border-radius:3px;margin-right:15px;flex-shrink:0}.list-container .item-info{flex-grow:1;text-align:left}.filename{font-size:.9em;font-weight:700;margin-top:5px;word-wrap:break-word}.dimensions{font-size:.8em;color:#777}.page-indicator{position:absolute;top:5px;right:5px;font-size:.7em;color:#fff;padding:2px 5px;border-radius:3px}.annotation-area{width:100%;box-sizing:border-box;margin-top:8px;padding:5px;border:1px dashed #ccc;border-radius:4px;font-family:sans-serif;resize:vertical;min-height:40px;font-size:14px;font-weight:700;color:#d32f2f}.folder-annotation{margin:10px 0;min-height:50px}.list-container .annotation-area{margin-left:15px}@media print{html,body{width:100%;height:100%;margin:0;padding:0}.controls{display:none}body{background-color:#fff}#printable-content{width:100%;margin:0;padding:0;box-shadow:none}.folder-header{background-color:#f0f0f0!important;-webkit-print-color-adjust:exact}.page-indicator{background-color:var(--bg-color)!important;-webkit-print-color-adjust:exact}.annotation-area{border:1px solid #eee;resize:none;background-color:#fdfdfd!important;-webkit-print-color-adjust:exact}.hide-on-print{display:none!important}}</style>"""
        js_script = """<script>let state={view:"grid",gridItemWidth:200,listThumbSize:80};function switchView(e){if(state.view===e)return;state.view=e,document.querySelectorAll(".content-container").forEach(t=>{t.classList.remove("grid-container","list-container"),t.classList.add(e+"-container")}),document.getElementById("btn-grid").classList.toggle("active","grid"===e),document.getElementById("btn-list").classList.toggle("active","list"===e)}function changeSize(e){"grid"===state.view?(state.gridItemWidth=Math.max(80,Math.min(600,state.gridItemWidth+40*e)),document.documentElement.style.setProperty("--grid-item-width",state.gridItemWidth+"px")):(state.listThumbSize=Math.max(40,Math.min(200,state.listThumbSize+20*e)),document.documentElement.style.setProperty("--list-thumbnail-size",state.listThumbSize+"px"))}function updatePrintStyle(){let e=document.querySelector('input[name="orientation"]:checked').value;document.getElementById("page-orientation-style").innerHTML=`@page { size: A4 ${e}; margin: 1.5cm; }`;let t=document.getElementById("printable-content");t.style.width="portrait"===e?"180mm":"267mm"}function filterFiles(){let e=document.getElementById("search-box").value.toLowerCase();document.querySelectorAll(".folder-container").forEach(t=>{let i=t.querySelectorAll(".item"),l=0,n=0;i.forEach(t=>{let i=t.querySelector(".filename").textContent.toLowerCase();i.includes(e)?(t.style.display="flex",l++,n+=parseFloat(t.dataset.area)):t.style.display="none"});let a=t.querySelector(".folder-header"),s=t.querySelector(".folder-stats"),d=a.dataset.originalFiles,o=a.dataset.originalPages,r=a.dataset.originalSqm;""===e.trim()?s.textContent=`File: ${d} | Pagine: ${o} | Area: ${r} m²`:s.textContent=`File: ${l} (di ${d}) | Area: ${n.toFixed(2)} m²`,t.style.display=l>0?"":"none"})}function prepareAndPrint(){document.querySelectorAll(".annotation-area").forEach(e=>{e.classList.toggle("hide-on-print",""===e.value.trim())}),window.print()}document.addEventListener("DOMContentLoaded",()=>{switchView("grid"),updatePrintStyle(),document.documentElement.style.setProperty("--grid-item-width",state.gridItemWidth+"px"),document.documentElement.style.setProperty("--list-thumbnail-size",state.listThumbSize+"px")});</script>"""
        body = """<body><div class="controls"><button onclick="prepareAndPrint()" class="print-button">Stampa Pagina</button><input type="text" id="search-box" onkeyup="filterFiles()" placeholder="Cerca per nome file..."><div class="control-group"><label>Orientamento:</label><input type="radio" id="portrait" name="orientation" value="portrait" checked onchange="updatePrintStyle()"><label for="portrait">Verticale</label><input type="radio" id="landscape" name="orientation" value="landscape" onchange="updatePrintStyle()"><label for="landscape">Orizzontale</label></div><div class="view-switcher control-group"><button id="btn-grid" onclick="switchView('grid')">Griglia</button><button id="btn-list" onclick="switchView('list')">Elenco</button></div><div class="size-controls control-group"><label>Dimensione:</label><button class="size-btn" onclick="changeSize(-1)">-</button><button class="size-btn" onclick="changeSize(1)">+</button></div></div><div id="printable-content">"""
        
        for folder, pages in sorted(grouped_pages.items()):
            display_folder = self._get_display_path(pages[0]['file_info']).replace('\\', '/')
            
            # Calcolo statistiche basato sulle pagine selezionate
            num_files = len(set(p['file_info']['filename'] for p in pages))
            total_pages = len(pages)
            total_sqm = sum(p['file_info']['pages_details'][p['page_num']]['width_cm'] * p['file_info']['pages_details'][p['page_num']]['height_cm'] for p in pages) / 10000
            
            body += f"""<div class="folder-container"><div class="folder-header" data-original-files="{num_files}" data-original-pages="{total_pages}" data-original-sqm="{total_sqm:.2f}"><span class="folder-stats">File: {num_files} | Pagine: {total_pages} | Area: {total_sqm:.2f} m²</span><span class="folder-path">{display_folder}</span></div><textarea class="annotation-area folder-annotation" placeholder="Aggiungi un'annotazione per questa cartella..."></textarea><div class="content-container grid-container">"""
            
            for page_index, page_data in enumerate(pages):
                item_data = page_data['file_info']
                page_num = page_data['page_num']
                full_path = os.path.join(item_data['path'], item_data['filename'])
                page_count = item_data.get('page_count', 1)
                current_color = file_colors[page_index % len(file_colors)]

                try:
                    page_details = item_data["pages_details"][page_num]
                    area_sqm = (page_details.get('width_cm', 0) * page_details.get('height_cm', 0)) / 10000

                    if item_data['type'] in ('PDF', 'AI'):
                        with fitz.open(full_path) as doc_pdf:
                            page = doc_pdf.load_page(page_num); pix = page.get_pixmap(dpi=150)
                            img_data = pix.tobytes("png")
                    else:
                        with Image.open(full_path) as img:
                            img.thumbnail((400, 400)); img_buffer = io.BytesIO()
                            img.save(img_buffer, format='PNG'); img_data = img_buffer.getvalue()
                    b64_img = base64.b64encode(img_data).decode('utf-8')
                    img_src = f"data:image/png;base64,{b64_img}"
                    
                    dpi_info = f"({item_data['dpi_str']})" if item_data.get('dpi_str') else ""
                    
                    body += f'<div class="item" data-area="{area_sqm}">'
                    if page_count > 1: body += f'<div class="page-indicator" style="background-color: {current_color}; --bg-color: {current_color};">Pag. {page_num + 1}/{page_count}</div>'
                    
                    body += f'''<img class="item-img" src="{img_src}" alt="Anteprima"><div class="item-info" style="flex-grow: 1; display: flex; flex-direction: column; justify-content: center;"><div class="filename">{html.escape(item_data["filename"])}</div><div class="dimensions">{html.escape(page_details["dimensions_cm"])} cm {html.escape(dpi_info)}</div></div><textarea class="annotation-area" placeholder="Annotazione..."></textarea></div>'''
                except Exception as e: print(f"Impossibile creare anteprima HTML per {full_path}, pagina {page_num + 1}: {e}")
            body += '</div></div>'
        
        body += '</div></body>'
        return f"<!DOCTYPE html><html><head><meta charset='UTF-8'><title>Anteprima Report</title>{css}{js_script}</head>{body}</html>"

    def _generate_html_table_with_totals(self, pages_to_export, include_headers_footers=True):
        if not pages_to_export: return ""
        total_area = 0
        html_string = '<table border="1" style="border-collapse: collapse; width: 100%; font-family: sans-serif;">'
        
        if include_headers_footers:
            html_string += '<thead style="background-color: #e0e0e0;"><tr>'
            headers = ["Nome File / Pagina", "Dimensioni (cm)", "Percorso", "Area (m²)"]
            for h in headers: html_string += f'<th style="padding: 5px; text-align: left;">{html.escape(h)}</th>'
            html_string += '</tr></thead>'

        html_string += '<tbody>'
        for page_data in pages_to_export:
            html_string += '<tr>'
            
            file_info = page_data['file_info']
            page_num = page_data['page_num']
            page_details = file_info["pages_details"][page_num]
            
            page_count = file_info.get('page_count', 1)
            filename_display = file_info['filename'] if page_count == 1 else f"   {file_info['filename']} (Pag. {page_num + 1})"
            dimensions_display = page_details['dimensions_cm']
            path_display = self._get_display_path(file_info).replace('\\', '/')

            width_cm, height_cm = page_details.get('width_cm', 0), page_details.get('height_cm', 0)
            area_sqm = (width_cm * height_cm) / 10000
            total_area += area_sqm
            
            cell_style = 'padding: 5px; text-align: left;'
            filename_style = 'padding: 5px; padding-left: 20px; text-align: left;' if page_count > 1 else cell_style

            html_string += f'<td style="{filename_style}">{html.escape(filename_display)}</td>'
            html_string += f'<td style="{cell_style}">{html.escape(dimensions_display)}</td>'
            html_string += f'<td style="{cell_style}">{html.escape(path_display)}</td>'
            html_string += f'<td style="{cell_style}">{area_sqm:.4f}</td></tr>'

        html_string += '</tbody>'
        
        if include_headers_footers:
            html_string += f'<tfoot><tr style="font-weight: bold; background-color: #f0f0f0;"><td style="padding: 5px;" colspan="3">Totale ({len(pages_to_export)} elementi)</td><td style="padding: 5px;">{total_area:.4f} m²</td></tr></tfoot>'
            
        html_string += '</table>'
        return html_string

    def _set_clipboard_html(self, html_fragment: str):
        if os.name != 'nt': raise NotImplementedError("La copia formattata è supportata solo su Windows.")
        user32 = ctypes.WinDLL('user32'); kernel32 = ctypes.WinDLL('kernel32')
        wintypes.HGLOBAL = wintypes.HANDLE
        user32.OpenClipboard.argtypes = [wintypes.HWND]; user32.OpenClipboard.restype = wintypes.BOOL
        user32.CloseClipboard.restype = wintypes.BOOL; user32.EmptyClipboard.restype = wintypes.BOOL
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HGLOBAL]; user32.SetClipboardData.restype = wintypes.HGLOBAL
        user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]; user32.RegisterClipboardFormatW.restype = wintypes.UINT
        kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]; kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]; kernel32.GlobalLock.restype = wintypes.LPVOID
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]; kernel32.GlobalUnlock.restype = wintypes.BOOL
        kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
        GMEM_MOVEABLE = 0x0002
        CF_HTML = user32.RegisterClipboardFormatW("HTML Format")
        header_template = "Version:0.9\r\nStartHTML:{{:0>9}}\r\nEndHTML:{{:0>9}}\r\nStartFragment:{{:0>9}}\r\nEndFragment:{{:0>9}}\r\n"
        html_body_template = "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body><!--StartFragment-->{}<!--EndFragment--></body></html>"
        html_body = html_body_template.format(html_fragment)
        header_placeholder_utf8 = header_template.format(0, 0, 0, 0).encode('utf-8')
        html_body_utf8 = html_body.encode('utf-8')
        start_html = len(header_placeholder_utf8)
        start_fragment = start_html + html_body.find("<!--StartFragment-->") + len("<!--StartFragment-->")
        end_fragment = start_html + html_body.find("<!--EndFragment-->")
        end_html = start_html + len(html_body_utf8)
        final_header_utf8 = header_template.format(start_html, end_html, start_fragment, end_fragment).encode('utf-8')
        clipboard_data = final_header_utf8 + html_body_utf8
        h_global_mem = None
        if not user32.OpenClipboard(None): raise ctypes.WinError()
        try:
            user32.EmptyClipboard()
            h_global_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(clipboard_data) + 1)
            if not h_global_mem: raise ctypes.WinError()
            p_global_mem = kernel32.GlobalLock(h_global_mem)
            if not p_global_mem: raise ctypes.WinError()
            try: ctypes.memmove(p_global_mem, clipboard_data, len(clipboard_data))
            finally: kernel32.GlobalUnlock(h_global_mem)
            if not user32.SetClipboardData(CF_HTML, h_global_mem): raise ctypes.WinError()
            h_global_mem = None
        finally:
            user32.CloseClipboard()
            if h_global_mem: kernel32.GlobalFree(h_global_mem)

    def copy_formatted_to_clipboard(self, selection_mode=False):
        pages_to_copy = self.get_pages_for_selection(selection_mode)
        if not pages_to_copy:
            self.status_text.set("Nessuna riga da copiare."); return
            
        html_table = self._generate_html_table_with_totals(pages_to_copy, include_headers_footers=True)
        if not html_table: return
        try:
            self._set_clipboard_html(html_table)
            self.status_text.set(f"Tabella formattata ({len(pages_to_copy)} righe) copiata negli appunti.")
        except Exception as e:
            self.status_text.set("Errore: Impossibile copiare la tabella formattata.")
            messagebox.showerror("Errore Appunti", f"Non è stato possibile copiare l'HTML negli appunti.\nDettagli: {e}", parent=self)
            traceback.print_exc()

    def export_to_pdf(self, selection_mode=False):
        pages_to_export = self.get_pages_for_selection(selection_mode)
        if not pages_to_export:
            messagebox.showinfo("Informazione", "Nessun dato da esportare.", parent=self); return

        self._lock_ui()
        options_dialog = ExportOptionsWindow(self)
        self.wait_window(options_dialog)
        options = options_dialog.result
        
        if not options:
            self.status_text.set("Esportazione annullata.")
            self._unlock_ui()
            return

        file_path = filedialog.asksaveasfilename(parent=self, defaultextension=".pdf", filetypes=[("File PDF", "*.pdf")], title="Salva report PDF")
        if not file_path:
            self.status_text.set("Esportazione PDF annullata.")
            self._unlock_ui()
            return

        self.status_text.set("Creazione del report PDF in corso...")
        self.update_idletasks()

        thread = threading.Thread(target=self._build_pdf_thread, args=(file_path, options, pages_to_export), daemon=True)
        thread.start()

    def _build_pdf_thread(self, file_path, options, pages_to_export):
        try:
            grouped_pages = defaultdict(list)
            for page in pages_to_export:
                grouped_pages[page['file_info']['scan_root']].append(page)
            
            page_size = landscape(A4) if options['orientation'] == 'landscape' else A4
            doc = SimpleDocTemplate(file_path, pagesize=page_size, topMargin=1.5*cm, bottomMargin=1.5*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
            
            styles = getSampleStyleSheet()
            filename_style = ParagraphStyle('file_style', parent=styles['Normal'], fontSize=8, alignment=1)
            dims_style = ParagraphStyle('dims_style', parent=styles['Normal'], fontSize=7, textColor=colors.darkgrey, alignment=1)
            folder_header_style = ParagraphStyle('folder_header', parent=styles['h2'], backColor=colors.lightblue, padding=4, textColor=colors.black)
            
            story = []
            num_columns = options['columns']
            col_width = (doc.width / num_columns) - (cm * 0.2 * (num_columns - 1))

            for folder, pages in sorted(grouped_pages.items()):
                display_folder = self._get_display_path(pages[0]['file_info']).replace('\\', '/')
                num_files = len(set(p['file_info']['filename'] for p in pages))
                total_pages = len(pages)
                total_sqm = sum(p['file_info']['pages_details'][p['page_num']]['width_cm'] * p['file_info']['pages_details'][p['page_num']]['height_cm'] for p in pages) / 10000
                
                stats_text = f"File: {num_files} | Pagine: {total_pages} | Area: {total_sqm:.2f} m²"
                story.append(Paragraph(display_folder, folder_header_style))
                story.append(Paragraph(stats_text, styles['Normal']))
                story.append(Spacer(1, 0.5*cm))

                grid_data = []; row = []
                for page_data in pages:
                    item_data = page_data['file_info']
                    page_num = page_data['page_num']
                    full_path = os.path.join(item_data['path'], item_data['filename'])
                    page_count = item_data.get('page_count', 1)
                    
                    cell_content = []
                    try:
                        if item_data['type'] in ('PDF', 'AI'):
                            with fitz.open(full_path) as doc_pdf:
                                page = doc_pdf.load_page(page_num); pix = page.get_pixmap(dpi=150)
                                img_data = pix.tobytes("png"); img_report = ReportLabImage(io.BytesIO(img_data), width=col_width*0.9, height=col_width*0.9, kind='proportional')
                        else:
                            with Image.open(full_path) as img:
                                img.thumbnail((400, 400)); img_buffer = io.BytesIO()
                                img.save(img_buffer, format='PNG'); img_buffer.seek(0)
                                img_report = ReportLabImage(img_buffer, width=col_width*0.9, height=col_width*0.9, kind='proportional')
                        cell_content.append(img_report)
                    except Exception as e: print(f"Impossibile creare anteprima PDF per {full_path}: {e}")
                    
                    page_info = f" (Pag. {page_num + 1}/{page_count})" if page_count > 1 else ""
                    cell_content.append(Paragraph(item_data['filename'] + page_info, filename_style))
                    
                    dpi_info = f"({item_data['dpi_str']})" if item_data.get('dpi_str') else ""
                    dims_text = item_data['pages_details'][page_num]['dimensions_cm'] + f" cm {dpi_info}"
                    cell_content.append(Paragraph(dims_text, dims_style))
                    
                    row.append(cell_content)
                    if len(row) == num_columns: grid_data.append(row); row = []
                if row:
                    row.extend([""] * (num_columns - len(row)))
                    grid_data.append(row)
                
                if grid_data:
                    grid_table = Table(grid_data, colWidths=[col_width] * num_columns)
                    grid_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('BOX', (0,0), (-1,-1), 1, colors.lightgrey), ('PADDING', (0,0), (-1,-1), 6)]))
                    story.append(grid_table)
                story.append(PageBreak())
            
            if story and isinstance(story[-1], PageBreak): story.pop()
            
            doc.build(story)
            self.after(0, self.on_pdf_success, file_path)
        except Exception as e:
            self.after(0, self.on_pdf_error, e)

    def on_pdf_success(self, file_path):
        self.status_text.set("Report PDF creato con successo.")
        self._unlock_ui()
        self.update()
        messagebox.showinfo("Successo", f"Report esportato con successo in:\n{file_path}", parent=self)

    def on_pdf_error(self, e):
        traceback.print_exc()
        self.status_text.set("Errore durante la creazione del PDF.")
        self._unlock_ui()
        self.update()
        messagebox.showerror("Errore", f"Impossibile creare il PDF.\nErrore: {e}", parent=self)

# --- Funzione di caricamento richiesta da winfile.py ---
def create_tab(tab_view):
    """
    Crea la scheda "Liste e anteprime", inizializza l'app e restituisce
    il nome della scheda e l'istanza dell'app per la gestione eventi.
    """
    tab_name = "Liste e anteprime"
    tab = tab_view.add(tab_name)
    
    app_instance = FileScannerApp(master=tab)
    
    # Restituisce il nome della scheda e l'istanza dell'app
    return tab_name, app_instance

