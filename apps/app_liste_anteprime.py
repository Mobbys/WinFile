# app_liste_anteprime.py - v5.0.16 (Fix definitivo layout Elenco A4)
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk, Menu
import os
import threading
from PIL import Image, ImageTk, ImageDraw
import fitz  # PyMuPDF
import io
from collections import defaultdict
import traceback
import webbrowser
import tempfile
import base64
import csv
import html
import functools
from concurrent.futures import ThreadPoolExecutor

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

DEFAULT_DPI = 96
PREVIEW_SIZE = (300, 300)


class ExportOptionsWindow(ctk.CTkToplevel):
    """
    Finestra di dialogo (convertita in CustomTkinter) per le opzioni di esportazione.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Opzioni di Esportazione")
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
    Classe principale dell'applicazione.
    """
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        self.status_text = ctk.StringVar(value="Trascina file o cartelle qui, oppure usa il pulsante.")
        self.preview_image = None
        self.scan_results = []
        self.is_scanning = False
        self.sort_state = {'col': None, 'reverse': False}
        self.bottom_button_groups = []

        self.create_widgets()
        self.create_context_menu()
        self.style_treeview()
        self.bottom_controls_frame.bind('<Configure>', self._rearrange_button_groups)

    def handle_drop(self, event):
        if self.is_scanning: return
        try:
            paths = self.tk.splitlist(event.data)
        except Exception:
            paths = []
        if paths: self.run_scan(paths)
        else: self.status_text.set("Nessun file o cartella valida trascinata.")

    def _rearrange_button_groups(self, event=None):
        if not self.winfo_viewable(): return
        available_width = self.bottom_controls_frame.winfo_width()
        padding_x, padding_y = 2, 2
        cursor_x, cursor_y, row_height = padding_x, padding_y, 0
        for group in self.bottom_button_groups:
            group_width, group_height = group.winfo_reqwidth(), group.winfo_reqheight()
            if cursor_x + group_width + padding_x > available_width and cursor_x > padding_x:
                cursor_y += row_height + padding_y
                cursor_x = padding_x
                row_height = 0
            group.place(x=cursor_x, y=cursor_y)
            cursor_x += group_width + padding_x
            if group_height > row_height: row_height = group_height
        required_height = cursor_y + row_height + padding_y
        if self.bottom_controls_frame.winfo_reqheight() != required_height:
             self.bottom_controls_frame.configure(height=required_height)

    def create_widgets(self):
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.select_button = ctk.CTkButton(top_frame, text="Aggiungi Cartella", command=self.select_folder_dialog)
        self.select_button.pack(side="left", padx=5, pady=5)
        self.clear_button = ctk.CTkButton(top_frame, text="Svuota Lista", command=self.clear_results, state="disabled")
        self.clear_button.pack(side="left", padx=5, pady=5)

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        main_frame.grid_columnconfigure(0, weight=3); main_frame.grid_columnconfigure(1, weight=1); main_frame.grid_rowconfigure(0, weight=1)

        tree_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        tree_frame.grid_rowconfigure(0, weight=1); tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("filename", "dimensions_cm", "area_sqm", "path")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings")
        self.tree.column("#0", width=30, stretch=False, anchor="center")
        self.tree.heading("#0", text="")
        self.tree.heading("filename", text="Nome File / Pagina", command=lambda: self.sort_by_column("filename"))
        self.tree.heading("dimensions_cm", text="Dimensioni (cm)", command=lambda: self.sort_by_column("dimensions_cm"))
        self.tree.heading("area_sqm", text="Area (m²)", command=lambda: self.sort_by_column("area_sqm"))
        self.tree.heading("path", text="Sottocartella", command=lambda: self.sort_by_column("path"))
        self.tree.column("filename", width=250); 
        self.tree.column("dimensions_cm", width=160, anchor="center")
        self.tree.column("area_sqm", width=110, anchor="center")
        self.tree.column("path", width=150)
        scrollbar = ctk.CTkScrollbar(tree_frame, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_item_select)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", self.open_selected_file)
        
        preview_frame = ctk.CTkFrame(main_frame)
        preview_frame.grid(row=0, column=1, sticky="nsew")
        self.preview_label = ctk.CTkLabel(preview_frame, text="")
        self.preview_label.pack(fill="both", expand=True, padx=5, pady=5)

        self.bottom_controls_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_controls_frame.pack(fill="x", padx=10, pady=5)

        copy_frame = ctk.CTkFrame(self.bottom_controls_frame)
        ctk.CTkLabel(copy_frame, text="Copia").pack(anchor="w", pady=(2, 2), padx=10)
        self.copy_all_button = ctk.CTkButton(copy_frame, text="Copia Tabella", command=self.copy_all_to_clipboard, state="disabled")
        self.copy_all_button.pack(side="left", padx=5, pady=(0, 5))
        self.copy_formatted_button = ctk.CTkButton(copy_frame, text="Copia Formattata", command=lambda: self.copy_formatted_to_clipboard(selection_mode=False), state="disabled")
        self.copy_formatted_button.pack(side="left", padx=5, pady=(0, 5))
        self.bottom_button_groups.append(copy_frame)

        export_frame = ctk.CTkFrame(self.bottom_controls_frame)
        ctk.CTkLabel(export_frame, text="Esporta / Stampa").pack(anchor="w", pady=(2, 2), padx=10)
        self.print_button = ctk.CTkButton(export_frame, text="Stampa Tabella", command=lambda: self.print_table(selection_mode=False), state="disabled")
        self.print_button.pack(side="left", padx=5, pady=(0, 5))
        self.export_html_button = ctk.CTkButton(export_frame, text="Anteprima Miniature", command=self.export_to_html, state="disabled")
        self.export_html_button.pack(side="left", padx=5, pady=(0, 5))
        self.export_csv_button = ctk.CTkButton(export_frame, text="Esporta in CSV", command=lambda: self.export_to_csv(selection_mode=False), state="disabled")
        self.export_csv_button.pack(side="left", padx=5, pady=(0, 5))
        self.export_pdf_button = ctk.CTkButton(export_frame, text="Esporta in PDF", command=self.export_to_pdf, state="disabled")
        self.export_pdf_button.pack(side="left", padx=5, pady=(0, 5))
        self.bottom_button_groups.append(export_frame)

        status_bar = ctk.CTkFrame(self, height=30)
        status_bar.pack(side="bottom", fill="x", padx=10, pady=(5, 10))
        status_label = ctk.CTkLabel(status_bar, textvariable=self.status_text, anchor="w")
        status_label.pack(side="left", padx=10)
    
    def sort_by_column(self, col):
        if not self.scan_results: return
        if self.sort_state['col'] == col: self.sort_state['reverse'] = not self.sort_state['reverse']
        else: self.sort_state['col'], self.sort_state['reverse'] = col, False
        
        if col == "filename": 
            sort_key = lambda item: item['filename'].lower()
        elif col == "path": 
            sort_key = lambda item: self._get_display_path(item).lower()
        elif col == "dimensions_cm":
            def get_area(item):
                try: 
                    page_detail = item['pages_details'][0]
                    return page_detail.get('trim_area_sqm', page_detail.get('area_sqm', 0))
                except (IndexError, KeyError): return 0
            sort_key = get_area
        elif col == "area_sqm":
            def get_area_sqm(item):
                try: 
                    page_detail = item['pages_details'][0]
                    return page_detail.get('trim_area_sqm', page_detail.get('area_sqm', 0))
                except (IndexError, KeyError): return 0
            sort_key = get_area_sqm
        else: return

        self.scan_results.sort(key=sort_key, reverse=self.sort_state['reverse'])
        self.update_column_headings()
        self.repopulate_treeview()

    def update_column_headings(self):
        for col in ("filename", "dimensions_cm", "area_sqm", "path"):
            original_text = self.tree.heading(col, "text").split(" ")[0]
            if col == "filename": original_text = "Nome File / Pagina"
            elif col == "dimensions_cm": original_text = "Dimensioni (cm)"
            elif col == "area_sqm": original_text = "Area (m²)"
            elif col == "path": original_text = "Sottocartella"
            
            if col == self.sort_state['col']:
                arrow = '▼' if self.sort_state['reverse'] else '▲'
                self.tree.heading(col, text=f"{original_text} {arrow}")
            else:
                 self.tree.heading(col, text=original_text)

    def _lock_ui(self):
        for button in [self.select_button, self.clear_button, self.copy_all_button, self.copy_formatted_button, self.print_button, self.export_html_button, self.export_pdf_button, self.export_csv_button]:
            button.configure(state="disabled")

    def _unlock_ui(self):
        self.select_button.configure(state="normal")
        state = "normal" if self.scan_results else "disabled"
        for button in [self.clear_button, self.copy_all_button, self.copy_formatted_button, self.print_button, self.export_html_button, self.export_pdf_button, self.export_csv_button]:
            button.configure(state=state)

    def create_context_menu(self):
        self.option_add("*Menu.font", ("Segoe UI", 10))
        self.context_menu = Menu(self, tearoff=0)
        self.context_menu.add_command(label="Copia selezione negli appunti", command=self.copy_selection_to_clipboard)
        self.context_menu.add_command(label="Copia selezione formattata", command=lambda: self.copy_formatted_to_clipboard(selection_mode=True))
        self.context_menu.add_command(label="Stampa selezione", command=lambda: self.print_table(selection_mode=True))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Esporta selezione in CSV", command=lambda: self.export_to_csv(selection_mode=True))
        self.context_menu.add_command(label="Crea Anteprima Miniature...", command=lambda: self.export_to_html(selection_mode=True))
        self.context_menu.add_command(label="Esporta selezione in PDF", command=lambda: self.export_to_pdf(selection_mode=True))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Rimuovi selezionati", command=self.remove_selected_items)

    def show_context_menu(self, event):
        if self.tree.selection(): self.context_menu.post(event.x_root, event.y_root)

    def style_treeview(self):
        style = ttk.Style()
        font, row_height = ("Segoe UI", 10), 22 
        if ctk.get_appearance_mode() == "Dark":
            bg, fg, field_bg = "#2B2B2B", "white", "#343638"
            sel, odd, even = "#1f6aa5", "#242424", "#2B2B2B"
            folder_bg = "#343638"
        else:
            bg, fg, field_bg = "white", "black", "#EAEAEA"
            sel, odd, even = "#3484d0", "#F7F7F7", "white"
            folder_bg = "#E0E0E0"
        style.theme_use("default")
        style.configure("Treeview", background=bg, foreground=fg, fieldbackground=field_bg, borderwidth=0, font=font, rowheight=row_height)
        style.map('Treeview', background=[('selected', sel)])
        style.configure("Treeview.Heading", background=field_bg, foreground=fg, relief="flat", font=(font[0], font[1], "bold"))
        style.map("Treeview.Heading", background=[('active', sel)])
        self.tree.tag_configure('oddrow', background=odd)
        self.tree.tag_configure('evenrow', background=even)
        self.tree.tag_configure('folder_row', background=folder_bg, font=(font[0], font[1], "bold"))

    def get_file_details(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            details = {"filename": os.path.basename(file_path), "type": ext.replace('.', '').upper(), "path": os.path.dirname(file_path)}
            if ext in ('.jpg', '.jpeg', '.tif', '.tiff', '.png'):
                with Image.open(file_path) as img:
                    w_px, h_px = img.size; dpi_x, dpi_y = img.info.get('dpi', (DEFAULT_DPI, DEFAULT_DPI))
                    w_cm = (w_px / (dpi_x or DEFAULT_DPI)) * 2.54; h_cm = (h_px / (dpi_y or DEFAULT_DPI)) * 2.54
                    area_sqm = (w_cm * h_cm) / 10000
                    details.update({"w_px": w_px, "h_px": h_px, "page_count": 1, "pages_details": [{"dimensions_cm": f"{w_cm:.2f} x {h_cm:.2f}", "width_cm": w_cm, "height_cm": h_cm, "area_sqm": area_sqm}], "dpi_str": f"{int(dpi_x or DEFAULT_DPI)} DPI"})
                    return details
            elif ext in ('.pdf', '.ai'):
                try:
                    with fitz.open(file_path) as doc:
                        if not doc: raise ValueError("Documento vuoto")
                        pages_details = []
                        for i, page in enumerate(doc):
                            rect = page.rect
                            w_cm, h_cm = (rect.width/72)*2.54, (rect.height/72)*2.54
                            area_sqm = (w_cm * h_cm) / 10000
                            page_detail = {
                                "dimensions_cm": f"{w_cm:.2f} x {h_cm:.2f}", 
                                "width_cm": w_cm, "height_cm": h_cm, 
                                "area_sqm": area_sqm
                            }
                            trim_rect = page.trimbox
                            if trim_rect and trim_rect != rect:
                                trim_w_cm = (trim_rect.width / 72) * 2.54
                                trim_h_cm = (trim_rect.height / 72) * 2.54
                                trim_area_sqm = (trim_w_cm * trim_h_cm) / 10000
                                page_detail['trim_dimensions_cm'] = f"{trim_w_cm:.2f} x {trim_h_cm:.2f}"
                                page_detail['trim_area_sqm'] = trim_area_sqm
                            pages_details.append(page_detail)
                        details.update({"page_count": len(doc), "pages_details": pages_details, "type": "AI" if ext == '.ai' else "PDF"})
                        return details
                except Exception:
                    details.update({"type": "AI (Non compatibile)" if ext == '.ai' else "PDF (Danneggiato)", "page_count": 1, "pages_details": [{"dimensions_cm": "Non rilevabili", "width_cm": 0, "height_cm": 0, "area_sqm": 0}]})
                    return details
            else:
                details.update({"page_count": 1, "pages_details": [{"dimensions_cm": "Anteprima non disponibile", "width_cm": 0, "height_cm": 0, "area_sqm": 0}], "type": "NON SUPPORTATO"})
                return details
        except Exception as e:
            print(f"Errore analisi file {file_path}: {e}")
            return {"filename": os.path.basename(file_path), "type": "ERRORE", "path": os.path.dirname(file_path), "page_count": 1, "pages_details": [{"dimensions_cm": "Errore lettura", "width_cm": 0, "height_cm": 0, "area_sqm": 0}]}

    def update_scan_progress(self, current_path, count):
        self.status_text.set(f"Scansione: {os.path.basename(current_path)}... ({count} file trovati)")

    def process_paths(self, paths):
        self._lock_ui()
        found_files = []
        for path in paths:
            scan_root = os.path.normpath(path)
            if os.path.isdir(scan_root):
                for root_dir, _, files in os.walk(scan_root):
                    self.after(0, self.update_scan_progress, root_dir, len(found_files))
                    for file in files:
                        full_path = os.path.join(root_dir, file)
                        if details := self.get_file_details(full_path):
                            details['scan_root'] = scan_root
                            found_files.append(details)
            elif os.path.isfile(scan_root):
                if details := self.get_file_details(scan_root):
                    details['scan_root'] = os.path.dirname(scan_root)
                    found_files.append(details)
        self.after(0, self.add_scan_results, found_files)

    def add_scan_results(self, new_files):
        self.scan_results.extend(new_files)
        unique_results, seen_paths = [], set()
        for item in self.scan_results:
            full_path = os.path.join(item['path'], item['filename'])
            if full_path not in seen_paths:
                unique_results.append(item); seen_paths.add(full_path)
        self.scan_results = unique_results
        if self.sort_state['col']: self.sort_by_column(self.sort_state['col'])
        else: self.repopulate_treeview()

    def _get_display_path(self, file_info):
        scan_root = file_info.get('scan_root', '')
        file_dir = file_info['path']
        
        if not scan_root or os.path.normpath(file_dir) == os.path.normpath(scan_root):
            return "."
        
        try:
            relative_path = os.path.relpath(file_dir, scan_root)
            return relative_path.replace(os.path.sep, ' > ')
        except ValueError:
            return os.path.basename(file_dir)

    def repopulate_treeview(self):
        selection = self.tree.selection()
        open_items = {item for item in self.tree.get_children() if self.tree.item(item, "open")}
        
        self.tree.delete(*self.tree.get_children())

        grouped_results = defaultdict(list)
        for item in self.scan_results:
            grouped_results[item.get('scan_root', 'N/A')].append(item)

        total_rows = 0
        sorted_scan_roots = sorted(grouped_results.keys())

        for i, scan_root in enumerate(sorted_scan_roots):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            folder_id = f"folder_{scan_root}"
            is_open = folder_id in open_items or not open_items
            
            total_folder_sqm = sum(p.get('area_sqm', 0) for file_info in grouped_results[scan_root] for p in file_info['pages_details'])
            total_folder_trim_sqm = sum(p.get('trim_area_sqm', p.get('area_sqm', 0)) for file_info in grouped_results[scan_root] for p in file_info['pages_details'])
            
            area_display = f"{total_folder_sqm:.4f}"
            if abs(total_folder_sqm - total_folder_trim_sqm) > 0.0001:
                area_display += f" ({total_folder_trim_sqm:.4f})"

            self.tree.insert("", "end", iid=folder_id, values=(os.path.basename(scan_root), f"({len(grouped_results[scan_root])} file)", area_display, scan_root),
                             open=is_open, tags=('folder_row',))
            
            for file_info in grouped_results[scan_root]:
                path_id = os.path.join(file_info['path'], file_info['filename'])
                display_path = self._get_display_path(file_info)

                if (page_count := file_info.get('page_count', 1)) > 1:
                    total_file_sqm = sum(p.get('area_sqm', 0) for p in file_info['pages_details'])
                    total_file_trim_sqm = sum(p.get('trim_area_sqm', p.get('area_sqm', 0)) for p in file_info['pages_details'])
                    
                    area_display_file = f"{total_file_sqm:.4f}"
                    if abs(total_file_sqm - total_file_trim_sqm) > 0.0001:
                         area_display_file += f" ({total_file_trim_sqm:.4f})"

                    parent_id = f"file_{path_id}"
                    self.tree.insert(folder_id, "end", iid=parent_id, values=(f"{file_info['filename']} ({page_count} pagine)", "Multi-pagina", area_display_file, display_path), 
                                     tags=(tag,), open=(parent_id in open_items))
                    for page_num in range(page_count):
                        page_details = file_info["pages_details"][page_num]
                        
                        dims_display = page_details["dimensions_cm"]
                        if 'trim_dimensions_cm' in page_details:
                            dims_display += f" ({page_details['trim_dimensions_cm']})"
                        
                        area_page_display = f"{page_details.get('area_sqm', 0):.4f}"
                        if 'trim_area_sqm' in page_details:
                             area_page_display += f" ({page_details.get('trim_area_sqm', 0):.4f})"

                        self.tree.insert(parent_id, "end", iid=f"{path_id}_{page_num}", values=(f"   Pagina {page_num + 1}", dims_display, area_page_display, ""), tags=(tag,))
                    total_rows += page_count
                else:
                    page_details = file_info["pages_details"][0]
                    
                    dims_display = page_details["dimensions_cm"]
                    if 'trim_dimensions_cm' in page_details:
                        dims_display += f" ({page_details['trim_dimensions_cm']})"
                        
                    area_page_display = f"{page_details.get('area_sqm', 0):.4f}"
                    if 'trim_area_sqm' in page_details:
                         area_page_display += f" ({page_details.get('trim_area_sqm', 0):.4f})"

                    self.tree.insert(folder_id, "end", iid=f"{path_id}_0", values=(file_info['filename'], dims_display, area_page_display, display_path), tags=(tag,))
                    total_rows += 1

        if selection:
            try: self.tree.selection_set(selection)
            except Exception: pass
            
        self.status_text.set(f"Scansione completata. Trovati {len(self.scan_results)} file ({total_rows} elementi) in {len(sorted_scan_roots)} cartelle.")
        if not selection: self.preview_label.configure(image=None)
        self._unlock_ui()
        self.is_scanning = False
    
    def _find_item_data_by_id(self, item_id):
        if not item_id or item_id.startswith("folder_"): return None, -1
        
        path_part, page_to_show = item_id.replace("file_", ""), 0
        if "_" in os.path.basename(path_part):
            try:
                base_path, page_str = path_part.rsplit('_', 1)
                page_to_show = int(page_str)
                path_part = base_path
            except ValueError: pass
            
        for data in self.scan_results:
            if os.path.join(data['path'], data['filename']) == path_part:
                return data, page_to_show
        return None, -1

    def select_folder_dialog(self):
        if self.is_scanning: return
        if folder := filedialog.askdirectory(parent=self): self.run_scan([folder])

    def run_scan(self, paths):
        self.is_scanning = True
        self.status_text.set("Avvio scansione...")
        threading.Thread(target=self.process_paths, args=(paths,), daemon=True).start()

    def clear_results(self):
        self.scan_results = []
        self.sort_state = {'col': None, 'reverse': False}
        self.update_column_headings()
        self.tree.delete(*self.tree.get_children())
        self.status_text.set("Lista svuotata. Pronto per una nuova scansione.")
        self.preview_label.configure(image=None)
        self._unlock_ui()

    def _create_placeholder_image(self, size, text):
        img = Image.new('RGB', size, color = (200, 200, 200))
        d = ImageDraw.Draw(img)
        try:
            font = ctk.CTkFont(family="Arial", size=20)
            d.text((10,10), text, fill=(0,0,0), font=font)
        except:
            d.text((10,10), text, fill=(0,0,0))
        return img

    def on_item_select(self, event):
        if not (sel := self.tree.selection()): return
        selected_data, page_to_show = self._find_item_data_by_id(sel[0])
        if not selected_data: 
            self.preview_label.configure(image=None)
            return
        full_path = os.path.join(selected_data['path'], selected_data['filename'])
        
        if selected_data['type'] == "NON SUPPORTATO":
            img = self._create_placeholder_image(PREVIEW_SIZE, "Anteprima non disponibile")
            self.preview_image = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.preview_label.configure(image=self.preview_image)
            return

        try:
            if selected_data['type'] in ('PDF', 'AI'):
                with fitz.open(full_path) as doc:
                    pix = doc.load_page(page_to_show).get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            else: 
                img = Image.open(full_path)

            img.thumbnail(PREVIEW_SIZE, Image.Resampling.LANCZOS)
            self.preview_image = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.preview_label.configure(image=self.preview_image)
        except Exception as e:
            self.status_text.set(f"Impossibile generare anteprima.")
            print(f"Errore anteprima: {e}")
            img = self._create_placeholder_image(PREVIEW_SIZE, "Anteprima non disponibile")
            self.preview_image = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.preview_label.configure(image=self.preview_image)

    def open_selected_file(self, event):
        if not (focus_id := self.tree.focus()): return
        if focus_id.startswith("folder_"):
             try:
                folder_path = self.tree.item(focus_id, "values")[2]
                if os.path.exists(folder_path): os.startfile(os.path.realpath(folder_path))
                else: messagebox.showerror("Errore", f"Cartella non trovata:\n{folder_path}", parent=self)
             except Exception as e: messagebox.showerror("Errore Apertura", f"Impossibile aprire la cartella.\n{e}", parent=self)
             return
        
        selected_data, _ = self._find_item_data_by_id(focus_id)
        if not selected_data: return print(f"Dati non validi: {focus_id}")
        try:
            full_path = os.path.join(selected_data['path'], selected_data['filename'])
            if os.path.exists(full_path): os.startfile(os.path.realpath(full_path))
            else: messagebox.showerror("Errore", f"File non trovato:\n{full_path}", parent=self)
        except Exception as e: messagebox.showerror("Errore Apertura", f"Impossibile aprire il file.\n{e}", parent=self)

    def remove_selected_items(self):
        if not (sel := self.tree.selection()): return
        paths_to_remove = set()
        folders_to_remove = set()
        
        for iid in sel:
            if iid.startswith("folder_"):
                folders_to_remove.add(self.tree.item(iid, "values")[2])
            else:
                data, _ = self._find_item_data_by_id(iid)
                if data:
                    paths_to_remove.add(os.path.join(data['path'], data['filename']))

        self.scan_results = [
            item for item in self.scan_results 
            if (os.path.join(item['path'], item['filename']) not in paths_to_remove and
                item.get('scan_root') not in folders_to_remove)
        ]
        self.repopulate_treeview()

    def get_pages_for_selection(self, selection_mode=False):
        if not selection_mode:
            return [{'file_info': fi, 'page_num': pn} for fi in self.scan_results for pn in range(fi.get('page_count', 1))]
        if not (sel := self.tree.selection()): return []
        
        selected_pages_set = set()
        for iid in sel:
            if iid.startswith("folder_"):
                folder_path = self.tree.item(iid, "values")[2]
                for item in self.scan_results:
                    if item.get('scan_root') == folder_path:
                         for pn in range(item.get('page_count', 1)):
                            selected_pages_set.add((item['path'], item['filename'], pn))
            else:
                data, page_num = self._find_item_data_by_id(iid)
                if data:
                    if iid.startswith("file_"): 
                         for pn in range(data.get('page_count', 1)):
                            selected_pages_set.add((data['path'], data['filename'], pn))
                    else: 
                        selected_pages_set.add((data['path'], data['filename'], page_num))

        all_pages = [{'file_info': fi, 'page_num': pn} for fi in self.scan_results for pn in range(fi.get('page_count', 1))]
        selected_pages = [p for p in all_pages if (p['file_info']['path'], p['file_info']['filename'], p['page_num']) in selected_pages_set]
        return selected_pages


    def export_to_html(self, selection_mode=False):
        pages_to_export = self.get_pages_for_selection(selection_mode)
        if not pages_to_export:
            return messagebox.showinfo("Informazione", "Nessun dato da esportare.", parent=self)
        
        quality = 'fast'

        self._lock_ui()
        self.status_text.set("Preparazione anteprima miniature...")
        self.update_idletasks()
        threading.Thread(target=self._build_html_thread, args=(pages_to_export, quality), daemon=True).start()

    def _build_html_thread(self, pages_to_export, quality):
        try:
            html_content = self._generate_html_content(pages_to_export, quality)
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.html', encoding='utf-8') as f:
                temp_file_path = f.name
                f.write(html_content)
            self.after(0, self.on_html_success, temp_file_path)
        except Exception as e: self.after(0, self.on_html_error, e)

    def on_html_success(self, file_path):
        try:
            if os.name == 'nt': os.startfile(os.path.realpath(file_path))
            else: webbrowser.open(f'file://{os.path.realpath(file_path)}')
            self.status_text.set("Anteprima miniature creata con successo.")
        except Exception as e:
            self.status_text.set("Errore apertura anteprima.")
            messagebox.showerror("Errore Apertura", f"Impossibile aprire il file HTML.\n{e}", parent=self)
        finally: self._unlock_ui(); self.update()

    def on_html_error(self, e):
        traceback.print_exc()
        self.status_text.set("Errore durante la creazione dell'anteprima.")
        self._unlock_ui(); self.update()
        messagebox.showerror("Errore", f"Impossibile creare l'anteprima.\n{e}", parent=self)

    def _generate_single_thumbnail(self, task_args):
        page_data, pdf_dpi, img_thumb_size = task_args
        item_data, page_num = page_data['file_info'], page_data['page_num']
        full_path = os.path.join(item_data['path'], item_data['filename'])

        try:
            img_data = b""
            if item_data['type'] == "NON SUPPORTATO":
                img = self._create_placeholder_image((200, 150), "Anteprima non disponibile")
                img_buffer = io.BytesIO()
                img.save(img_buffer, format='PNG')
                img_data = img_buffer.getvalue()
            elif item_data['type'] in ('PDF', 'AI'):
                with fitz.open(full_path) as doc_pdf:
                    pix = doc_pdf.load_page(page_num).get_pixmap(dpi=pdf_dpi)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img.thumbnail(img_thumb_size, Image.Resampling.LANCZOS)
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format='PNG')
                    img_data = img_buffer.getvalue()
            else:
                with Image.open(full_path) as img:
                    img.thumbnail(img_thumb_size, Image.Resampling.LANCZOS)
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format='PNG')
                    img_data = img_buffer.getvalue()
            
            return f"data:image/png;base64,{base64.b64encode(img_data).decode('utf-8')}"
        except Exception as e:
            print(f"Errore anteprima per {full_path}, pag {page_num + 1}: {e}")
            img = self._create_placeholder_image((200, 150), "Anteprima non disponibile")
            img_buffer = io.BytesIO()
            img.save(img_buffer, format='PNG')
            img_data = img_buffer.getvalue()
            return f"data:image/png;base64,{base64.b64encode(img_data).decode('utf-8')}"

    def _generate_html_content(self, pages_to_export, quality):
        grouped_pages = defaultdict(list)
        selected_pages_keys = {(os.path.join(p['file_info']['path'], p['file_info']['filename']), p['page_num']) for p in pages_to_export}
        
        for page in pages_to_export:
             full_path = os.path.join(page['file_info']['path'], page['file_info']['filename'])
             if (full_path, page['page_num']) in selected_pages_keys:
                grouped_pages[page['file_info']['scan_root']].append(page)

        file_colors = ['#DB4437', '#4285F4', '#F4B400', '#0F9D58', '#AB47BC', '#E91E63', '#9C27B0', '#673AB7', '#009688']
        unique_file_paths = sorted({os.path.join(p['file_info']['path'], p['file_info']['filename']) for p in pages_to_export})
        file_path_to_color = {path: file_colors[i % len(file_colors)] for i, path in enumerate(unique_file_paths)}
        
        js_libraries = """
        <script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        """
        
        if quality == 'fast':
            pdf_dpi = 72
            img_thumb_size = (750, 750)
        else: # high quality
            pdf_dpi = 150
            img_thumb_size = (1500, 1500)

        css = """<style>
            @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');
            :root { --a4-width: 21cm; --a4-height: 29.7cm; --item-width: 200px; }
            body{font-family:'Roboto',sans-serif;margin:0;background-color:#d2d2d2; color:#333;}
            .controls{display:flex;flex-wrap:wrap;align-items:center;gap:10px;background:#fff;padding:10px 20px;box-shadow:0 2px 4px #0000001a;position:sticky;top:0;z-index:1000}
            .controls button, .controls input, .controls select {padding:8px 12px;font-size:14px;border:1px solid #ccc;border-radius:5px;font-family:'Roboto',sans-serif;}
            .controls button {background-color:#e9e9e9;color:#333;cursor:pointer;font-weight:700}
            .print-button{background-color:#4285f4;color:#fff;border-color:#4285f4}
            .control-group{display:flex;align-items:center;gap:5px;padding:5px;border:1px solid #e0e0e0;border-radius:5px}
            #view-container { padding: 20px; }
            .page {
                background: white;
                box-shadow: 0 0 10px rgba(0,0,0,0.2);
                margin: 20px auto;
                position: relative;
                width: var(--a4-width);
                height: var(--a4-height);
            }
            .page.landscape { width: var(--a4-height); height: var(--a4-width); }
            .page-content { padding: 1cm; display: flex; flex-wrap: wrap; align-content: flex-start; gap: 5px; height: calc(100% - 2cm); width: calc(100% - 2cm); overflow: hidden; }
            
            .folder-header{
                display: flex;
                justify-content: space-between;
                align-items: center;
                background-color:#e0e0e0;
                padding:10px 15px;
                border-left:5px solid #4285f4;
                border-radius:0 5px 5px 0;
                width:100%;
                box-sizing:border-box;
                margin-bottom:5px;
            }
            .folder-path-input {
                font-family: 'Roboto', sans-serif;
                font-weight: 700;
                font-size: 1.2em;
                color: #333;
                border: none;
                background: transparent;
                padding: 2px 5px;
                margin: -2px -5px; /* Counteract padding to keep alignment */
                width: 70%;
                border-radius: 3px;
            }
            .folder-path-input:focus {
                outline: none;
                background: #fff;
                box-shadow: 0 0 0 2px #4285f4;
            }
            .folder-stats{font-size:.9em;color:#555;line-height:1.5em; text-align: right;}

            .subfolder-separator-container { width: 100%; display: flex; align-items: center; gap: 10px; margin: 15px 0 5px 0; }
            .subfolder-separator { flex-grow: 1; border: 0; border-top: 1px solid #ccc; }
            .breadcrumb-container { display: flex; align-items: center; gap: 5px; }
            .breadcrumb-crumb { padding: 2px 8px; border-radius: 4px; color: white; font-size: 0.9em; font-weight: bold; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
            .breadcrumb-separator { font-weight: bold; color: #555; }
            .subfolder-stats { font-size: 0.8em; color: #666; white-space: nowrap; }
            
            .page-content.page-layout .folder-header, .page-content.page-layout .folder-annotation, .page-content.page-layout .subfolder-separator-container { flex-basis: 100%; }
            .page-content.list-layout .folder-header, .page-content.list-layout .folder-annotation, .page-content.list-layout .subfolder-separator-container { width: 100%; }
            .page-content.grid-layout .folder-header, .page-content.grid-layout .folder-annotation, .page-content.grid-layout .subfolder-separator-container { grid-column: 1 / -1; }
            .folder-annotation { margin-bottom: 10px; }
            
            .page-content.grid-layout { display: grid; grid-template-columns: repeat(auto-fill, minmax(var(--item-width), 1fr)); gap: 10px; align-content: start; }
            .page-content.list-layout { display: flex; flex-direction: column; gap: 8px; flex-wrap: nowrap; }
            .page-content.page-layout { display: flex; flex-wrap: wrap; gap: 5px; justify-content: center; }
            
            .item {box-sizing: border-box; background-color:#fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); cursor: grab; display: flex; flex-direction: column; position: relative;}
            .item.sortable-ghost {opacity: 0.4;}
            .item-img-container { width: 100%; background-color: #fff; display: flex; align-items: center; justify-content: center; border-bottom:1px solid #ccc; }
            .item-img{display:block; max-width: 100%; max-height: 100%; object-fit: contain; border: 1px solid #000; box-sizing: border-box;}
            
            .item-info {
                padding: 8px; font-size: 12px; text-align: left; background: #f8f9fa; 
                width: 100%; box-sizing: border-box;
                display: flex; flex-direction: column;
                flex-grow: 1;
            }
            .item-info-header { 
                display: flex; justify-content: flex-end;
                width: 100%;
                order: 3;
                margin-top: 4px;
            }
            .metadata { font-size: 11px; color: #6c757d; order: 1; }
            .filename-container {
                margin-top: 5px;
                border-top: 1px solid #eee;
                padding-top: 5px;
                order: 2;
            }
            .filename { font-weight: 700; color: #212529; word-break: break-word; }
            .trim-info{font-size: 10px; color: #d9534f; font-weight: bold;}
            .page-indicator{font-size:.8em;color:#fff;padding:2px 5px;border-radius:3px; text-shadow: 1px 1px 2px #000; flex-shrink: 0; font-weight: bold;}

            .item-checkbox {position: absolute; top: 8px; left: 8px; width: 20px; height: 20px; z-index: 10; cursor: pointer; background-color: rgba(255,255,255,0.7); border-radius: 3px;}
            
            .page-content.page-layout .item { width: var(--item-width); }
            .page-content.grid-layout .item-img-container { height: calc(var(--item-width) * 0.75); }
            .page-content.list-layout .item { flex-direction: row; align-items: stretch; width: 100%; }
            .page-content.list-layout .item-img-container { width: calc(var(--item-width)); max-height: 300px; flex-shrink: 0; border-right: 1px solid #ccc; border-bottom: none; padding: 5px; box-sizing: border-box; }
            .page-content.list-layout .item-info { flex-grow: 1; border: none; justify-content: center; }
            
            .annotation-area{width:100% !important;box-sizing:border-box;margin:0;padding:5px;border:1px dashed #ccc;border-top: 1px solid #ccc;font-family:sans-serif;resize:vertical;min-height:40px;font-size:12px; color: red;}
            #loader {position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.7); color: white; display: flex; align-items: center; justify-content: center; font-size: 2em; z-index: 2000;}
            
            .summary-container { 
                font-size: 12px;
                width: 100%; 
                box-sizing: border-box;
            }
            .summary-container h2 { margin-top: 0; border-bottom: 2px solid #ccc; padding-bottom: 8px; font-size: 1.8em; }
            .summary-list-wrapper {
                column-count: 1;
                column-gap: 2em;
            }
            .summary-list-wrapper.two-columns {
                column-count: 2;
            }
            .summary-container ul { list-style-type: none; padding-left: 0; margin: 0; }
            .summary-container > ul > li { margin-bottom: 15px; font-size: 1.3em; break-inside: avoid; }
            .summary-container ul ul { padding-left: 20px; margin-top: 6px; }
            .summary-container ul ul li { font-weight: bold; font-size: 0.9em; }
            .summary-container ul ul ul { list-style-type: '- '; padding-left: 20px; }
            .summary-container ul ul ul li { padding: 3px 0; font-size: 1.1em; color: #333; font-weight: normal;}
            .summary-container ul ul ul li small { color: #555; font-size: 0.9em; }

            body.annotations-hidden .annotation-area { display: none !important; }
            body.trim-hidden .trim-info { display: none !important; }
            body.normal-hidden .normal-info { display: none !important; }
            .switch { position: relative; display: inline-block; width: 40px; height: 20px; vertical-align: middle;}
            .switch input { opacity: 0; width: 0; height: 0; }
            .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 20px; }
            .slider:before { position: absolute; content: ""; height: 16px; width: 16px; left: 2px; bottom: 2px; background-color: white; transition: .4s; border-radius: 50%; }
            input:checked + .slider { background-color: #f44336; }
            input:checked + .slider:before { transform: translateX(20px); }

            @media print{
                body{background-color:#fff !important; margin:0; padding:0;}
                .controls, #loader, .item-checkbox {display:none !important;}
                #view-container { padding: 0 !important; }
                .page { margin: 0; box-shadow: none; page-break-after: always; }
                .page:last-child { page-break-after: auto; }
                .folder-header{background-color:#f0f0f0!important;-webkit-print-color-adjust:exact;}
                .item.hide-for-print { display: none !important; }
                .item-info {background-color: #fff !important; border-color: #ddd !important; color: #000 !important; -webkit-print-color-adjust: exact;}
                .page-indicator{background-color:var(--bg-color)!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
            }
            </style>"""
        js_script = """<script>
            const state = {
                view: 'grid', 
                itemWidth: 200, 
                pageBreakOnSubfolder: false,
            };
            
            const pageSizes = {
                'A4': { width: '21cm', height: '29.7cm' },
                'A3': { width: '29.7cm', height: '42cm' }
            };

            function syncStateToSource() {
                const viewContainer = document.getElementById('view-container');
                const sourceData = document.getElementById('source-data');
                if (!viewContainer || !sourceData) return;

                // Sync inputs by updating the value attribute
                viewContainer.querySelectorAll('input.folder-path-input').forEach(visibleInput => {
                    const sourceInput = sourceData.querySelector(`#${CSS.escape(visibleInput.id)}`);
                    if (sourceInput) {
                        sourceInput.setAttribute('value', visibleInput.value);
                    }
                });

                // Sync textareas by updating their textContent, which is copied by cloneNode
                viewContainer.querySelectorAll('textarea.folder-annotation, textarea.annotation-area').forEach(visibleTextarea => {
                    const sourceTextarea = sourceData.querySelector(`#${CSS.escape(visibleTextarea.id)}`);
                    if (sourceTextarea) {
                        sourceTextarea.textContent = visibleTextarea.value;
                    }
                });
            }

            function updatePageLayout() {
                const size = document.getElementById('page-size-select').value;
                const dimensions = pageSizes[size];
                
                document.documentElement.style.setProperty('--a4-width', dimensions.width);
                document.documentElement.style.setProperty('--a4-height', dimensions.height);
                
                renderAllPages(); 
            }

            function switchView(viewName) {
                state.view = viewName;
                document.querySelectorAll('.view-switch-btn').forEach(btn => btn.style.backgroundColor = '');
                document.querySelector(`[onclick="switchView('${viewName}')"]`).style.backgroundColor = '#c0c0c0';
                renderAllPages();
            }

            function toggleAnnotationsVisibility(hide) {
                document.body.classList.toggle('annotations-hidden', hide);
            }

            function toggleTrimVisibility(hide) {
                document.body.classList.toggle('trim-hidden', hide);
            }

            function toggleNormalVisibility(hide) {
                document.body.classList.toggle('normal-hidden', hide);
            }

            function togglePageBreak(enabled) {
                state.pageBreakOnSubfolder = enabled;
                renderAllPages();
            }

            function generateSummaryHTML(sourceNode) {
                let listHtml = '<ul>';
                const folders = sourceNode.querySelectorAll('.folder-container');
                
                folders.forEach(folder => {
                    if (folder.style.display === 'none') return;
                    
                    const folderNameInput = folder.querySelector('.folder-path-input');
                    const folderName = folderNameInput ? folderNameInput.getAttribute('value') : 'Cartella Sconosciuta';
                    listHtml += `<li><strong>${folderName}</strong>`;

                    const subfolderSections = new Map();
                    let currentSectionName = 'File nella cartella principale';
                    let currentSectionItems = [];

                    Array.from(folder.children).forEach(child => {
                        if (child.classList.contains('subfolder-separator-container')) {
                            if (currentSectionItems.length > 0) {
                                subfolderSections.set(currentSectionName, currentSectionItems);
                            }
                            const breadcrumbs = Array.from(child.querySelectorAll('.breadcrumb-crumb')).map(c => c.textContent).join(' > ');
                            currentSectionName = breadcrumbs || 'Sottocartella';
                            if (currentSectionName === 'File nella cartella principale') {
                                const rootBreadcrumb = child.querySelector('.breadcrumb-crumb');
                                if (rootBreadcrumb) currentSectionName = rootBreadcrumb.textContent;
                            }
                            currentSectionItems = [];
                        } else if (child.classList.contains('item') && child.style.display !== 'none') {
                            currentSectionItems.push(child);
                        }
                    });
                    if (currentSectionItems.length > 0) {
                        subfolderSections.set(currentSectionName, currentSectionItems);
                    }

                    if (subfolderSections.size > 0) {
                        listHtml += '<ul>';
                        subfolderSections.forEach((items, name) => {
                             const visibleItems = items.filter(item => item.style.display !== 'none');
                             if(visibleItems.length > 0) {
                                listHtml += `<li>${name}<ul>`;
                                visibleItems.forEach(item => {
                                    const filename = item.querySelector('.filename')?.textContent || '';
                                    const pageIndicator = item.querySelector('.page-indicator')?.textContent || '';
                                    listHtml += `<li>${filename} ${pageIndicator}</li>`;
                                });
                                listHtml += '</ul></li>';
                             }
                        });
                        listHtml += '</ul>';
                    }
                    listHtml += '</li>';
                });

                listHtml += '</ul>';
                
                let html = `<div class="summary-container">
                                <h2>Riepilogo Contenuto</h2>
                                <div class="summary-list-wrapper">${listHtml}</div>
                            </div>`;
                return html;
            }

            function renderAllPages() {
                syncStateToSource();
                const viewContainer = document.getElementById('view-container');
                viewContainer.innerHTML = '';
                
                const sourceData = document.getElementById('source-data');
                const contentToLayout = sourceData.cloneNode(true);
                const folders = Array.from(contentToLayout.querySelectorAll('.folder-container'));

                if (folders.length === 0) return;

                let currentPage, contentWrapper;
                let pageContentHeight = 0; 

                const setupNewPage = () => {
                    currentPage = createNewPage();
                    viewContainer.appendChild(currentPage);
                    contentWrapper = currentPage.querySelector('.page-content');
                    contentWrapper.className = `page-content`; // Default to no special layout
                    if (pageContentHeight === 0) {
                        pageContentHeight = contentWrapper.clientHeight;
                    }
                };
                
                const hasSubfolders = Array.from(sourceData.querySelectorAll('.subfolder-separator-container')).some(el => {
                    const parentFolder = el.closest('.folder-container');
                    return parentFolder && parentFolder.style.display !== 'none';
                });

                if (state.pageBreakOnSubfolder && hasSubfolders) {
                    setupNewPage();

                    const firstVisibleFolder = Array.from(folders).find(f => f.style.display !== 'none');
                    if (firstVisibleFolder) {
                        const firstFolderHeader = firstVisibleFolder.querySelector('.folder-header').cloneNode(true);
                        const input = firstFolderHeader.querySelector('.folder-path-input');
                        if(input) {
                            const sourceInput = document.getElementById(input.id);
                            input.value = sourceInput ? sourceInput.getAttribute('value') : '';
                        }
                        contentWrapper.appendChild(firstFolderHeader);
                    }
                    
                    const summaryDiv = document.createElement('div');
                    summaryDiv.innerHTML = generateSummaryHTML(contentToLayout);
                    
                    const summaryWrapper = summaryDiv.querySelector('.summary-list-wrapper');
                    
                    contentWrapper.appendChild(summaryDiv);
                    
                    if (summaryWrapper.scrollHeight > pageContentHeight) {
                        summaryWrapper.classList.add('two-columns');
                    }
                    
                    // Paginate the summary itself if it still overflows
                    while (contentWrapper.scrollHeight > pageContentHeight) {
                        const allLIs = Array.from(contentWrapper.querySelectorAll('.summary-container li'));
                        let lastVisibleLI = null;

                        for(let i = allLIs.length - 1; i >= 0; i--) {
                            const li = allLIs[i];
                            if (li.offsetTop < pageContentHeight) {
                                lastVisibleLI = li;
                                break;
                            }
                        }

                        if (!lastVisibleLI) { 
                            break;
                        }

                        let breakList = lastVisibleLI.parentElement;
                        const itemsToMove = [];
                        let sibling = lastVisibleLI.nextElementSibling;
                        while(sibling) {
                            itemsToMove.push(sibling);
                            sibling = sibling.nextElementSibling;
                        }
                        
                        setupNewPage();
                        contentWrapper.className = 'page-content'; 
                        const newList = breakList.cloneNode(false);
                        itemsToMove.forEach(item => {
                            newList.appendChild(item);
                        });

                        let currentNewList = newList;
                        let currentOldList = breakList;
                        while(currentOldList.parentElement && currentOldList.parentElement.closest('.summary-container')) {
                             const newParentList = currentOldList.parentElement.cloneNode(false);
                             newParentList.appendChild(currentNewList);
                             currentNewList = newParentList;
                             
                             let parentSibling = currentOldList.parentElement.nextElementSibling;
                             while(parentSibling){
                                currentNewList.appendChild(parentSibling);
                                parentSibling = parentSibling.nextElementSibling;
                             }
                             currentOldList = currentOldList.parentElement;
                        }
                        contentWrapper.appendChild(currentNewList);
                    }


                    const hasVisibleItems = contentToLayout.querySelector(".item:not([style*='display: none'])");
                    if (hasVisibleItems) {
                        setupNewPage(); 
                    } else {
                        contentWrapper = null; 
                    }
                } else {
                    setupNewPage(); 
                }
                
                if (!contentWrapper) {
                    initDragAndDrop();
                    return;
                }
                
                contentWrapper.className = `page-content ${state.view}-layout`;

                for (const folder of folders) {
                    if(folder.style.display === 'none') continue;
                    
                    const childrenToProcess = Array.from(folder.children);
                    const isFirstFolder = folder === folders.find(f => f.style.display !== 'none');

                    for (const child of childrenToProcess) {
                        if(child.style.display === 'none') continue;
                        
                        if(state.pageBreakOnSubfolder && hasSubfolders) {
                            if (isFirstFolder && child.classList.contains('folder-header')) {
                                continue;
                            }
                            if (!isFirstFolder && child.classList.contains('folder-header') && contentWrapper.children.length > 0) {
                                setupNewPage();
                                contentWrapper.className = `page-content ${state.view}-layout`;
                            }
                            if (child.classList.contains('subfolder-separator-container') && contentWrapper.querySelector('.item')) {
                               setupNewPage();
                               contentWrapper.className = `page-content ${state.view}-layout`;
                            }
                        }

                        if(child.classList.contains('folder-header')){
                            const input = child.querySelector('.folder-path-input');
                            if(input) {
                                const sourceInput = document.getElementById(input.id);
                                input.value = sourceInput ? sourceInput.getAttribute('value') : '';
                            }
                        }
                        contentWrapper.appendChild(child);
                        if (pageContentHeight > 0 && contentWrapper.scrollHeight > pageContentHeight) {
                            contentWrapper.removeChild(child);
                            setupNewPage();
                            contentWrapper.className = `page-content ${state.view}-layout`;
                            contentWrapper.appendChild(child);
                        }
                    }
                }
                
                if (contentWrapper && contentWrapper.children.length === 0 && viewContainer.children.length > 1) {
                    viewContainer.removeChild(currentPage);
                }

                initDragAndDrop();
            }
            
            function createNewPage() {
                const page = document.createElement('div');
                page.className = 'page';
                const orientation = document.getElementById('page-orientation-select').value;
                if (orientation === 'landscape') {
                    page.classList.add('landscape');
                }
                page.innerHTML = '<div class="page-content"></div>';
                return page;
            }

            function changeSize(amount) {
                state.itemWidth = Math.max(100, Math.min(1600, state.itemWidth + amount * 25));
                document.getElementById('zoom-slider').value = state.itemWidth;
                document.documentElement.style.setProperty('--item-width', state.itemWidth + 'px');
                renderAllPages();
            }

            function setZoom(value) {
                state.itemWidth = parseInt(value, 10);
                document.documentElement.style.setProperty('--item-width', state.itemWidth + 'px');
                renderAllPages();
            }

            function filterFiles(){
                const query = document.getElementById("search-box").value.toLowerCase();
                document.querySelectorAll("#source-data .folder-container").forEach(folder => {
                    let visibleItems = 0;
                    folder.querySelectorAll(".item").forEach(item => {
                        const filename = item.querySelector(".filename").textContent.toLowerCase();
                        const isVisible = filename.includes(query);
                        item.style.display = isVisible ? "flex" : "none";
                        if(isVisible) visibleItems++;
                    });
                    folder.style.display = visibleItems > 0 ? "block" : "none";
                });
                renderAllPages();
            }
            
            function initDragAndDrop() {
                 document.querySelectorAll('.page-content').forEach(container => {
                    if (container.sortableInstance) container.sortableInstance.destroy();
                    container.sortableInstance = new Sortable(container, {
                        group: 'shared-items',
                        animation: 150,
                        ghostClass: 'sortable-ghost',
                        onEnd: (evt) => {
                           const itemId = evt.item.id;
                           const sourceItem = document.querySelector(`#source-data #${itemId}`);
                           const nextSibling = evt.item.nextElementSibling;
                           
                           if(nextSibling) {
                               const nextSourceSibling = document.querySelector(`#source-data #${nextSibling.id}`);
                               if (nextSourceSibling) {
                                   nextSourceSibling.parentNode.insertBefore(sourceItem, nextSourceSibling);
                               }
                           } else {
                               const toContainerId = evt.to.parentElement.id;
                               const fromContainerId = evt.from.parentElement.id;
                               // This is a simplified logic, might need adjustment for complex cases
                               const sourceParentContainer = document.querySelector(`#source-data #${evt.to.querySelector('.item').id}`).parentElement;
                               if(sourceParentContainer) {
                                    sourceParentContainer.appendChild(sourceItem);
                               }
                           }
                           renderAllPages();
                        }
                    });
                });
            }

            function toggleSelectAll(shouldBeChecked) {
                document.querySelectorAll('.item-checkbox').forEach(cb => { cb.checked = shouldBeChecked; });
            }

            function prepareAndPrintStandard() {
                document.querySelectorAll('#view-container .item').forEach(item => {
                    const checkbox = item.querySelector('.item-checkbox');
                    item.classList.toggle('hide-for-print', checkbox && !checkbox.checked);
                });

                const emptyAnnotations = [];
                if (!document.body.classList.contains('annotations-hidden')) {
                    document.querySelectorAll('#view-container .annotation-area').forEach(area => {
                        if (area.value.trim() === '') {
                            area.style.display = 'none';
                            emptyAnnotations.push(area);
                        }
                    });
                }

                window.print();

                emptyAnnotations.forEach(area => {
                    area.style.display = '';
                });
            }
            
            async function createWysiwygPdf(outputAction = 'save') {
                const { jsPDF } = window.jspdf;
                const loader = document.getElementById('loader');
                
                let printWindow = null;
                if (outputAction === 'print') {
                    printWindow = window.open('', '_blank');
                    if (!printWindow) { alert('Impossibile aprire la finestra di anteprima. Disabilitare il blocco pop-up e riprovare.'); return; }
                    printWindow.document.write('<html><head><title>Stampa Fedele</title></head><body><p>Generazione PDF...</p></body></html>');
                }
                
                loader.style.display = 'flex';
                
                try {
                    const orientation = document.getElementById('page-orientation-select').value;
                    const size = document.getElementById('page-size-select').value.toLowerCase();
                    const pdf = new jsPDF(orientation, 'mm', size);

                    const sourceElements = document.querySelectorAll('#view-container .page');
                    let firstPage = true;

                    for (const element of sourceElements) {
                        const contentToPrint = element.cloneNode(true);
                        
                        contentToPrint.querySelectorAll('input.folder-path-input').forEach(input => {
                            const newSpan = document.createElement('span');
                            newSpan.className = 'folder-path-input'; // Keep class for styling
                            newSpan.textContent = input.value;
                            input.parentNode.replaceChild(newSpan, input);
                        });

                        contentToPrint.querySelectorAll('.item:not([style*="display: none"])').forEach(clonedItem => {
                             const originalCheckbox = document.getElementById(clonedItem.id)?.querySelector('.item-checkbox');
                             if (originalCheckbox && !originalCheckbox.checked) {
                                 clonedItem.style.display = 'none';
                             } else {
                                clonedItem.querySelector('.item-checkbox')?.remove();
                                const originalTextarea = document.getElementById(clonedItem.id)?.querySelector('.annotation-area');
                                const clonedTextarea = clonedItem.querySelector('.annotation-area');
                                if (document.body.classList.contains('annotations-hidden') || (originalTextarea && originalTextarea.value.trim() === '')) {
                                    clonedTextarea?.remove();
                                } else if (originalTextarea && clonedTextarea) {
                                    clonedTextarea.textContent = originalTextarea.value;
                                }
                             }
                        });

                        contentToPrint.querySelectorAll('.folder-annotation, .subfolder-separator-container').forEach(clonedAnnotation => {
                           const originalAnnotation = document.getElementById(clonedAnnotation.id);
                           if (originalAnnotation && originalAnnotation.tagName === 'TEXTAREA') {
                               if (document.body.classList.contains('annotations-hidden') || (originalAnnotation && originalAnnotation.value.trim() === '')) {
                                   clonedAnnotation.remove();
                               } else if (originalAnnotation) {
                                   clonedAnnotation.textContent = originalAnnotation.value;
                               }
                           }
                        });
                        
                        document.body.appendChild(contentToPrint);
                        const canvas = await html2canvas(contentToPrint, { scale: 2.5 });
                        document.body.removeChild(contentToPrint);
                        
                        const imgData = canvas.toDataURL('image/png');
                        const imgProps = pdf.getImageProperties(imgData);
                        const pdfWidth = pdf.internal.pageSize.getWidth();
                        const pdfHeight = pdf.internal.pageSize.getHeight();
                        const imgWidth = pdfWidth;
                        const imgHeight = (imgProps.height * imgWidth) / imgProps.width;

                        if (!firstPage) pdf.addPage(size, orientation);
                        pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
                        firstPage = false;
                    }
                    
                    const pdfData = pdf.output('blob');
                    const url = URL.createObjectURL(pdfData);

                    if (outputAction === 'save') {
                        const link = document.createElement('a');
                        link.href = url;
                        link.download = 'Anteprima-Miniature.pdf';
                        document.body.appendChild(link); link.click(); document.body.removeChild(link);
                        URL.revokeObjectURL(url);
                    } else { 
                        printWindow.location.href = url;
                    }
                } catch (e) {
                    console.error("Errore generazione PDF:", e);
                    alert("Si è verificato un errore durante la generazione del PDF.");
                    if(printWindow) printWindow.close();
                } finally {
                    loader.style.display = 'none';
                }
            }
            
            function syncInputValues(source, dest) {
                const sourceInputs = source.querySelectorAll('input.folder-path-input');
                const destInputs = dest.querySelectorAll('input.folder-path-input');
                sourceInputs.forEach((sIn, i) => {
                    if (destInputs[i]) {
                        destInputs[i].value = sIn.value;
                    }
                });
            }

            document.addEventListener("DOMContentLoaded", () => {
                document.getElementById('zoom-slider').value = state.itemWidth;
                updatePageLayout(); // Chiamata iniziale
                switchView('grid');
            });
        </script>"""
        
        source_html = ""
        sorted_grouped_pages = sorted(grouped_pages.items(), key=lambda item: item[0])
        item_counter = 0

        for folder_idx, (folder, pages) in enumerate(sorted_grouped_pages):
            display_folder = os.path.basename(folder)
            num_files = len({(p['file_info']['path'], p['file_info']['filename']) for p in pages})
            total_pages = len(pages)
            
            total_sqm = sum(p['file_info']['pages_details'][p['page_num']]['area_sqm'] for p in pages)
            total_trim_sqm = sum(p['file_info']['pages_details'][p['page_num']].get('trim_area_sqm', p['file_info']['pages_details'][p['page_num']]['area_sqm']) for p in pages)
            
            folder_stats_text = f"File: {num_files} | Pagine: {total_pages} | Area: {total_sqm:.2f} m²"
            if abs(total_sqm - total_trim_sqm) > 0.0001:
                folder_stats_text += f" (Al vivo: {total_trim_sqm:.2f} m²)"

            folder_id = f"folder-{folder_idx}"
            input_id = f"folder-path-input-{folder_idx}"
            source_html += f"""<div class="folder-container" data-folder-id="{folder_id}">
                <div class="folder-header">
                    <input type="text" class="folder-path-input" id="{input_id}" value="{html.escape(display_folder)}">
                    <span class="folder-stats">{folder_stats_text}</span>
                </div>
                <textarea class="annotation-area folder-annotation" id="anno-{folder_id}" placeholder="Annotazione cartella..."></textarea>"""
            
            section_stats = defaultdict(lambda: {'files': set(), 'pages': 0, 'sqm': 0, 'trim_sqm': 0})
            for page_data in pages:
                section_key = self._get_display_path(page_data['file_info'])
                stats = section_stats[section_key]
                stats['files'].add(os.path.join(page_data['file_info']['path'], page_data['file_info']['filename']))
                stats['pages'] += 1
                page_details = page_data['file_info']['pages_details'][page_data['page_num']]
                stats['sqm'] += page_details.get('area_sqm', 0)
                stats['trim_sqm'] += page_details.get('trim_area_sqm', page_details.get('area_sqm', 0))
            
            pages.sort(key=lambda p: (self._get_display_path(p['file_info']) != ".", self._get_display_path(p['file_info'])))
            
            tasks = [(p, pdf_dpi, img_thumb_size) for p in pages]
            with ThreadPoolExecutor() as executor:
                b64_images = list(executor.map(self._generate_single_thumbnail, tasks))

            last_subfolder = None
            for i, page_data in enumerate(pages):
                current_subfolder = self._get_display_path(page_data['file_info'])
                
                if current_subfolder != last_subfolder:
                    stats = section_stats[current_subfolder]
                    num_files_sec = len(stats['files'])
                    total_pages_sec = stats['pages']
                    total_sqm_sec = stats['sqm']
                    total_trim_sqm_sec = stats['trim_sqm']
                    
                    stats_text = f"File: {num_files_sec} | Pagine: {total_pages_sec} | Area: {total_sqm_sec:.2f} m²"
                    if abs(total_sqm_sec - total_trim_sqm_sec) > 0.0001:
                        stats_text += f" (Al vivo: {total_trim_sqm_sec:.2f} m²)"

                    header_content = ''
                    if current_subfolder == ".":
                        header_content = '<div class="breadcrumb-container"><span class="breadcrumb-crumb" style="background-color: #6c757d;">File nella cartella principale</span></div>'
                    else:
                        header_content = '<div class="breadcrumb-container">'
                        crumbs = current_subfolder.split(' > ')
                        breadcrumb_colors = ['#4A90E2', '#50E3C2', '#F5A623', '#BD10E0', '#9013FE']
                        for j, crumb in enumerate(crumbs):
                            color = breadcrumb_colors[j % len(breadcrumb_colors)]
                            header_content += f'<span class="breadcrumb-crumb" style="background-color: {color};">{html.escape(crumb)}</span>'
                            if j < len(crumbs) - 1:
                                header_content += '<span class="breadcrumb-separator">&gt;</span>'
                        header_content += '</div>'

                    source_html += f'''<div class="subfolder-separator-container">
                        {header_content}
                        <hr class="subfolder-separator">
                        <span class="subfolder-stats">{stats_text}</span>
                    </div>'''
                    last_subfolder = current_subfolder

                item_data, page_num = page_data['file_info'], page_data['page_num']
                full_path = os.path.join(item_data['path'], item_data['filename'])
                page_details = item_data["pages_details"][page_num]
                
                img_src = b64_images[i]
                
                page_count = item_data.get('page_count', 1)
                current_color = file_path_to_color.get(full_path, '#808080')
                
                area_sqm_val = page_details.get('area_sqm', 0)
                trim_area_sqm_val = page_details.get('trim_area_sqm', area_sqm_val)
                
                dims_html = f'<div class="normal-info"><span>{page_details["dimensions_cm"]} cm &nbsp; {area_sqm_val:.3f} m²</span></div>'
                trim_html = ''
                if 'trim_dimensions_cm' in page_details:
                    trim_html = f'<div class="trim-info">Al vivo: {page_details["trim_dimensions_cm"]} cm &nbsp; {trim_area_sqm_val:.3f} m²</div>'
                
                page_indicator_span = f'<span class="page-indicator" style="--bg-color: {current_color}; background-color: {current_color};">Pag. {page_num + 1}/{page_count}</span>' if page_count > 1 else ''
                
                source_html += f"""<div class="item" id="item-{item_counter}" data-folder-id="{folder_id}">
                    <input type="checkbox" class="item-checkbox" checked>
                    <div class="item-img-container"><img class="item-img" src="{img_src}" alt="Anteprima"></div>
                    <div class="item-info">
                        <div class="metadata">{dims_html}{trim_html}</div>
                        <div class="filename-container">
                            <span class="filename">{html.escape(item_data["filename"])}</span>
                        </div>
                        <div class="item-info-header">{page_indicator_span}</div>
                    </div>
                    <textarea class="annotation-area" id="item-anno-{item_counter}" placeholder="Annotazione..."></textarea>
                </div>"""
                item_counter += 1
            source_html += '</div>'

        body = f"""<body>
            <div id="loader" style="display: none;"><span>Generazione PDF...</span></div>
            <div class="controls">
                <div class="control-group">
                    <button class="view-switch-btn" onclick="switchView('grid')">Griglia A4</button>
                    <button class="view-switch-btn" onclick="switchView('page')">Vista Nesting</button>
                    <button class="view-switch-btn" onclick="switchView('list')">Elenco A4</button>
                </div>
                <div class="control-group">
                    <label>Formato:</label>
                    <select id="page-size-select" onchange="updatePageLayout()">
                        <option value="A4">A4</option>
                        <option value="A3">A3</option>
                    </select>
                    <select id="page-orientation-select" onchange="updatePageLayout()">
                        <option value="portrait">Verticale</option>
                        <option value="landscape">Orizzontale</option>
                    </select>
                </div>
                <div class="control-group">
                    <button onclick="prepareAndPrintStandard()" title="Usa la stampa veloce del browser per la vista attiva.">Stampa Veloce</button>
                    <button onclick="createWysiwygPdf('print')" class="print-button" title="Genera un PDF fedele della vista attiva e la invia alla stampante.">Stampa Fedele</button>
                    <button onclick="createWysiwygPdf('save')" title="Genera e salva un PDF fedele della vista attiva.">Salva PDF Fedele</button>
                </div>
                <div class="control-group">
                    <button onclick="toggleSelectAll(true)">Seleziona Tutti</button>
                    <button onclick="toggleSelectAll(false)">Deseleziona Tutti</button>
                </div>
                <div class="control-group">
                    <input type="text" id="search-box" onkeyup="filterFiles()" placeholder="Cerca per nome file...">
                    <label>Dimensione:</label>
                    <button onclick="changeSize(-1)" title="Rimpicciolisci">-</button>
                    <input type="range" id="zoom-slider" min="100" max="1600" value="200" oninput="setZoom(this.value)">
                    <button onclick="changeSize(1)" title="Ingrandisci">+</button>
                </div>
                <div class="control-group" style="margin-left: auto; display: flex; gap: 15px;">
                    <div style="display: flex; align-items: center; gap: 5px;">
                        <label for="toggle-normal-cb" style="cursor:pointer; user-select: none;">Nascondi misure normali</label>
                        <label class="switch">
                            <input type="checkbox" id="toggle-normal-cb" onchange="toggleNormalVisibility(this.checked)">
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div style="display: flex; align-items: center; gap: 5px;">
                        <label for="toggle-trim-cb" style="cursor:pointer; user-select: none;">Nascondi misure al vivo</label>
                        <label class="switch">
                            <input type="checkbox" id="toggle-trim-cb" onchange="toggleTrimVisibility(this.checked)">
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div style="display: flex; align-items: center; gap: 5px;">
                        <label for="toggle-annotations-cb" style="cursor:pointer; user-select: none;">Nascondi Annotazioni</label>
                        <label class="switch">
                            <input type="checkbox" id="toggle-annotations-cb" onchange="toggleAnnotationsVisibility(this.checked)">
                            <span class="slider"></span>
                        </label>
                    </div>
                    <div style="display: flex; align-items: center; gap: 5px;">
                        <label for="toggle-page-break-cb" style="cursor:pointer; user-select: none;">Interruzione per sottocartella</label>
                        <label class="switch">
                            <input type="checkbox" id="toggle-page-break-cb" onchange="togglePageBreak(this.checked)">
                            <span class="slider"></span>
                        </label>
                    </div>
                </div>
            </div>
            <div id="view-container"></div>
            <div id="source-data" style="display:none;">{source_html}</div>
        </body>"""
        
        return f"<!DOCTYPE html><html><head><meta charset='UTF-8'><title>Anteprima Miniature</title>{js_libraries}{css}{js_script}</head>{body}</html>"


    def copy_all_to_clipboard(self):
        if not self.scan_results: return
        header = ["Nome File / Pagina", "Dimensioni (cm)", "Sottocartella"]
        lines = ["\t".join(header)]
        
        grouped_results = defaultdict(list)
        for item in self.scan_results:
            grouped_results[item.get('scan_root', 'N/A')].append(item)

        for folder, items in sorted(grouped_results.items()):
            lines.append(f"\n--- {os.path.basename(folder)} ---")

            subfolder_grouped_items = defaultdict(list)
            for item in items:
                subfolder_grouped_items[self._get_display_path(item)].append(item)
            
            sorted_subfolders = sorted(subfolder_grouped_items.keys(), key=lambda k: (k == ".", k))

            for subfolder in sorted_subfolders:
                subfolder_items = subfolder_grouped_items[subfolder]
                
                if subfolder != ".":
                    sub_num_files = len({os.path.join(item['path'], item['filename']) for item in subfolder_items})
                    sub_total_pages = sum(item.get('page_count', 1) for item in subfolder_items)
                    sub_total_sqm = sum(pd.get('area_sqm', 0) for item in subfolder_items for pd in item['pages_details'])
                    sub_total_trim_sqm = sum(pd.get('trim_area_sqm', pd.get('area_sqm', 0)) for item in subfolder_items for pd in item['pages_details'])
                    
                    area_text = f"{sub_total_trim_sqm:.2f} m²"
                    if abs(sub_total_sqm - sub_total_trim_sqm) > 0.0001:
                        area_text += f" (Orig: {sub_total_sqm:.2f} m²)"

                    stats_line = f"  --- {subfolder} (File: {sub_num_files}, Pagine: {sub_total_pages}, Area: {area_text}) ---"
                    lines.append(stats_line)

                for file_info in subfolder_items:
                    rel_path = self._get_display_path(file_info)
                    page_count = file_info.get('page_count', 1)
                    prefix = "    " if subfolder != "." else ""

                    if page_count > 1:
                        lines.append(f"{prefix}{file_info['filename']}\tMulti-pagina\t{rel_path}")
                        for i in range(page_count):
                            page_details = file_info['pages_details'][i]
                            dims_display = page_details['dimensions_cm']
                            if 'trim_dimensions_cm' in page_details:
                                dims_display += f" ({page_details['trim_dimensions_cm']})"
                            lines.append(f"{prefix}  Pagina {i+1}\t{dims_display}\t")
                    else:
                        page_details = file_info['pages_details'][0]
                        dims_display = page_details['dimensions_cm']
                        if 'trim_dimensions_cm' in page_details:
                            dims_display += f" ({page_details['trim_dimensions_cm']})"
                        lines.append(f"{prefix}{file_info['filename']}\t{dims_display}\t{rel_path}")

        self.clipboard_clear(); self.clipboard_append("\n".join(lines))
        self.status_text.set("Tabella raggruppata copiata negli appunti.")
        
    def copy_selection_to_clipboard(self):
        selected_pages = self.get_pages_for_selection(selection_mode=True)
        if not selected_pages: return
        
        header = ["Nome File / Pagina", "Dimensioni (cm)", "Sottocartella"]
        lines = ["\t".join(header)]
        
        grouped_pages = defaultdict(list)
        for page_data in selected_pages:
            grouped_pages[page_data['file_info'].get('scan_root', 'N/A')].append(page_data)

        for folder, pages in sorted(grouped_pages.items()):
            lines.append(f"\n--- {os.path.basename(folder)} ---")
            
            subfolder_grouped_pages = defaultdict(list)
            for page_data in pages:
                subfolder_grouped_pages[self._get_display_path(page_data['file_info'])].append(page_data)

            sorted_subfolders = sorted(subfolder_grouped_pages.keys(), key=lambda k: (k == ".", k))

            for subfolder in sorted_subfolders:
                subfolder_pages = subfolder_grouped_pages[subfolder]

                if subfolder != ".":
                    sub_num_files = len({(p['file_info']['path'], p['file_info']['filename']) for p in subfolder_pages})
                    sub_total_pages = len(subfolder_pages)
                    sub_total_sqm = sum(p['file_info']['pages_details'][p['page_num']].get('area_sqm', 0) for p in subfolder_pages)
                    sub_total_trim_sqm = sum(p['file_info']['pages_details'][p['page_num']].get('trim_area_sqm', p['file_info']['pages_details'][p['page_num']].get('area_sqm', 0)) for p in subfolder_pages)
                    
                    area_text = f"{sub_total_trim_sqm:.2f} m²"
                    if abs(sub_total_sqm - sub_total_trim_sqm) > 0.0001:
                        area_text += f" (Orig: {sub_total_sqm:.2f} m²)"
                    
                    stats_line = f"  --- {subfolder} (File: {sub_num_files}, Pagine: {sub_total_pages}, Area: {area_text}) ---"
                    lines.append(stats_line)

                for page_data in subfolder_pages:
                    item, page_num = page_data['file_info'], page_data['page_num']
                    rel_path = self._get_display_path(item)
                    page_details = item["pages_details"][page_num]
                    page_count = item.get('page_count', 1)
                    prefix = "    " if subfolder != "." else ""
                    
                    filename_display = f"{item['filename']} (Pag. {page_num + 1})" if page_count > 1 else item['filename']
                    dims_display = page_details['dimensions_cm']
                    if 'trim_dimensions_cm' in page_details:
                        dims_display += f" ({page_details['trim_dimensions_cm']})"

                    lines.append(f"{prefix}{filename_display}\t{dims_display}\t{rel_path}")

        self.clipboard_clear(); self.clipboard_append("\n".join(lines))
        self.status_text.set(f"Selezione raggruppata copiata ({len(selected_pages)} righe).")

    def print_table(self, selection_mode=False):
        pages_to_print = self.get_pages_for_selection(selection_mode)
        if not pages_to_print: return messagebox.showinfo("Informazione", "Nessuna riga da stampare.", parent=self)
        html_string = self._generate_html_table_with_totals(pages_to_print)
        html_content = f"""<!DOCTYPE html><html><head><meta charset='UTF-8'><title>Stampa Tabella</title><style>@media print{{@page{{margin:1.5cm}}body{{font-family:sans-serif;-webkit-print-color-adjust:exact}}table{{width:100%;border-collapse:collapse}}th,td{{border:1px solid black;padding:5px;text-align:left}}th{{background-color:#e0e0e0!important}}tfoot{{display:none;}}}}</style></head><body onload="window.print()">{html_string}</body></html>"""
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.html', encoding='utf-8') as f:
                f.write(html_content)
            webbrowser.open(f'file://{os.path.realpath(f.name)}')
            self.status_text.set("Apertura anteprima di stampa...")
        except Exception as e:
            traceback.print_exc(); self.status_text.set("Errore preparazione stampa.")
            messagebox.showerror("Errore Stampa", f"Impossibile aprire l'anteprima.\n{e}", parent=self)

    def export_to_csv(self, selection_mode=False):
        pages_to_export = self.get_pages_for_selection(selection_mode)
        if not pages_to_export: return messagebox.showinfo("Informazione", "Nessun dato da esportare.", parent=self)
        self._lock_ui()
        file_path = filedialog.asksaveasfilename(parent=self, defaultextension=".csv", filetypes=[("File CSV", "*.csv")], title="Salva lista come CSV")
        if not file_path: self._unlock_ui(); return self.status_text.set("Esportazione CSV annullata.")
        self.status_text.set("Creazione CSV in corso..."); self.update_idletasks()
        threading.Thread(target=self._build_csv_thread, args=(file_path, pages_to_export), daemon=True).start()

    def _build_csv_thread(self, file_path, pages_to_export):
        try:
            pages_to_export.sort(key=lambda p: p['file_info'].get('scan_root', 'N/A'))
            
            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow([
                    "Nome File", "Tipo", "Pagina", 
                    "Dimensioni (cm)", "Area (m²)", 
                    "Dimensioni Al vivo (cm)", "Area Al vivo (m²)",
                    "Sottocartella", "Cartella Principale"
                ])
                
                last_folder = None
                for page_data in pages_to_export:
                    item, page_num = page_data['file_info'], page_data['page_num']
                    
                    current_folder = item.get('scan_root', 'N/A')
                    if current_folder != last_folder:
                        if last_folder is not None:
                            writer.writerow([])
                        last_folder = current_folder
                        
                    display_path = self._get_display_path(item)
                    page_details = item['pages_details'][page_num]
                    
                    area_sqm = page_details.get('area_sqm', 0)
                    trim_dims = page_details.get('trim_dimensions_cm', '')
                    trim_area = page_details.get('trim_area_sqm', 0)
                    page_str = f"{page_num + 1} di {item.get('page_count', 1)}"

                    writer.writerow([
                        item['filename'], item['type'], page_str, 
                        page_details['dimensions_cm'], f"{area_sqm:.4f}",
                        trim_dims, f"{trim_area:.4f}" if trim_area > 0 else "",
                        display_path, item.get('scan_root', '')
                    ])
            
            self.after(0, self.on_csv_success, file_path)
        except Exception as e: self.after(0, self.on_csv_error, e)

    def on_csv_success(self, file_path):
        self.status_text.set("File CSV creato con successo.")
        self._unlock_ui(); self.update()
        messagebox.showinfo("Successo", f"Lista esportata con successo in:\n{file_path}", parent=self)
        
    def on_csv_error(self, e):
        traceback.print_exc(); self.status_text.set("Errore creazione CSV.")
        self._unlock_ui(); self.update()
        messagebox.showerror("Errore", f"Impossibile creare il file CSV.\n{e}", parent=self)

    def _generate_html_table_with_totals(self, pages_to_export, include_headers_footers=True):
        if not pages_to_export: return ""
        
        html_string = '<table border="1" style="border-collapse: collapse; width: 100%; font-family: sans-serif;">'
        if include_headers_footers:
            html_string += '<thead style="background-color: #e0e0e0;"><tr>'
            for h in ["Nome File / Pagina", "Dimensioni (cm)", "Sottocartella", "Area (m²)"]:
                html_string += f'<th style="padding: 5px; text-align: left;">{html.escape(h)}</th>'
            html_string += '</tr></thead>'
        
        html_string += '<tbody>'
        
        grouped_pages = defaultdict(list)
        for page_data in pages_to_export:
            grouped_pages[page_data['file_info'].get('scan_root', 'N/A')].append(page_data)

        total_area = 0
        total_trim_area = 0
        has_trim_box = False

        for folder, pages in sorted(grouped_pages.items()):
            folder_html = f'<tr style="background-color: #f0f0f0; font-weight: bold;"><td colspan="4" style="padding: 5px; border-left: 3px solid #4285f4;">{html.escape(os.path.basename(folder))}</td></tr>'
            html_string += folder_html

            subfolder_grouped_pages = defaultdict(list)
            for page_data in pages:
                subfolder_grouped_pages[self._get_display_path(page_data['file_info'])].append(page_data)

            sorted_subfolders = sorted(subfolder_grouped_pages.keys(), key=lambda k: (k == ".", k))

            for subfolder in sorted_subfolders:
                subfolder_pages = subfolder_grouped_pages[subfolder]

                if subfolder != ".":
                    sub_num_files = len({(p['file_info']['path'], p['file_info']['filename']) for p in subfolder_pages})
                    sub_total_pages = len(subfolder_pages)
                    sub_total_area = sum(p['file_info']['pages_details'][p['page_num']].get('area_sqm', 0) for p in subfolder_pages)
                    sub_total_trim_area = sum(p['file_info']['pages_details'][p['page_num']].get('trim_area_sqm', p['file_info']['pages_details'][p['page_num']].get('area_sqm', 0)) for p in subfolder_pages)
                    
                    sub_area_text = f"{sub_total_trim_area:.4f} m²"
                    if abs(sub_total_area - sub_total_trim_area) > 0.0001:
                        sub_area_text += f" (Orig: {sub_total_area:.4f} m²)"

                    stats_text = f"File: {sub_num_files} | Pagine: {sub_total_pages} | Area: {sub_area_text}"
                    subfolder_header_html = f'<tr style="background-color: #fafafa;"><td colspan="4" style="padding: 4px 10px; font-weight: bold; color: #333; border-bottom: 1px solid #ddd; border-top: 1px solid #ddd;">{html.escape(subfolder)}: <span style="font-weight:normal; color:#555">{stats_text}</span></td></tr>'
                    html_string += subfolder_header_html
                
                for page_data in subfolder_pages:
                    file_info, page_num = page_data['file_info'], page_data['page_num']
                    page_details = file_info["pages_details"][page_num]
                    page_count = file_info.get('page_count', 1)
                    filename_display = f"{file_info['filename']} (Pag. {page_num + 1})" if page_count > 1 else file_info['filename']
                    
                    dims_display = page_details["dimensions_cm"]
                    if 'trim_dimensions_cm' in page_details:
                        dims_display += f"<br><small style='color:red;'>Al vivo: {page_details['trim_dimensions_cm']}</small>"
                        has_trim_box = True

                    area_sqm = page_details.get('area_sqm', 0)
                    trim_area_sqm = page_details.get('trim_area_sqm', area_sqm)
                    total_area += area_sqm
                    total_trim_area += trim_area_sqm
                    
                    area_display = f"{area_sqm:.4f}"
                    if 'trim_area_sqm' in page_details:
                        area_display = f"{trim_area_sqm:.4f}"

                    path_display = self._get_display_path(file_info)
                    html_string += f'<tr><td style="padding:5px">{html.escape(filename_display)}</td><td style="padding:5px">{dims_display}</td><td style="padding:5px">{html.escape(path_display)}</td><td style="padding:5px">{area_display}</td></tr>'
                    
        html_string += '</tbody>'
        html_string += '</table>'
        
        if include_headers_footers:
            total_text = f"Totale ({len(pages_to_export)} elementi)"
            area_text = f"{total_trim_area:.4f} m²"
            if has_trim_box and abs(total_area - total_trim_area) > 0.0001:
                area_text += f"<br><small>(Originale: {total_area:.4f} m²)</small>"

            html_string += f'''
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px; page-break-inside: avoid;">
                <tr style="font-weight: bold; background-color: #f0f0f0;">
                    <td style="padding: 5px; border: 1px solid black;" colspan="3">{total_text}</td>
                    <td style="padding: 5px; border: 1px solid black;">{area_text}</td>
                </tr>
            </table>
            '''
        return html_string

    def _set_clipboard_html(self, html_fragment: str):
        if os.name != 'nt':
            raise NotImplementedError("La copia formattata è supportata solo su Windows.")

        user32 = ctypes.WinDLL('user32')
        kernel32 = ctypes.WinDLL('kernel32')
        user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
        user32.RegisterClipboardFormatW.restype = wintypes.UINT
        user32.OpenClipboard.argtypes = [wintypes.HWND]
        user32.OpenClipboard.restype = wintypes.BOOL
        user32.EmptyClipboard.restype = wintypes.BOOL
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
        user32.SetClipboardData.restype = wintypes.HANDLE
        user32.CloseClipboard.restype = wintypes.BOOL
        kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = wintypes.LPVOID
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalUnlock.restype = wintypes.BOOL
        kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalFree.restype = wintypes.HGLOBAL

        CF_HTML = user32.RegisterClipboardFormatW("HTML Format")
        if CF_HTML == 0:
            raise OSError("Impossibile registrare il formato HTML per gli appunti.")

        header = (
            "Version:0.9\r\n"
            "StartHTML:{:010d}\r\n"
            "EndHTML:{:010d}\r\n"
            "StartFragment:{:010d}\r\n"
            "EndFragment:{:010d}\r\n"
        )
        html_body = f"<!DOCTYPE html><html><head><style>td{{vertical-align:top;}}</style></head><body><!--StartFragment-->{html_fragment}<!--EndFragment--></body></html>"
        
        header_utf8 = header.encode('utf-8')
        html_body_utf8 = html_body.encode('utf-8')

        start_html = len(header_utf8)
        start_fragment = start_html + html_body.find("<!--StartFragment-->") + len("<!--StartFragment-->")
        end_fragment = start_html + html_body.find("<!--EndFragment-->")
        end_html = start_html + len(html_body_utf8)

        final_header = header.format(start_html, end_html, start_fragment, end_fragment).encode('utf-8')
        clipboard_data = final_header + html_body_utf8
        
        GMEM_MOVEABLE = 0x0002
        h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(clipboard_data) + 1)
        if not h_mem:
            raise MemoryError("GlobalAlloc ha fallito.")

        p_mem = kernel32.GlobalLock(h_mem)
        if not p_mem:
            kernel32.GlobalFree(h_mem)
            raise MemoryError("GlobalLock ha fallito.")

        try:
            ctypes.memmove(p_mem, clipboard_data, len(clipboard_data))
        finally:
            kernel32.GlobalUnlock(h_mem)

        if not user32.OpenClipboard(self.winfo_id()):
            kernel32.GlobalFree(h_mem)
            raise OSError("Impossibile aprire gli appunti.")

        try:
            user32.EmptyClipboard()
            if not user32.SetClipboardData(CF_HTML, h_mem):
                kernel32.GlobalFree(h_mem)
                raise OSError("Impossibile impostare i dati negli appunti.")
        finally:
            user32.CloseClipboard()

    def copy_formatted_to_clipboard(self, selection_mode=False):
        pages_to_copy = self.get_pages_for_selection(selection_mode)
        if not pages_to_copy: return self.status_text.set("Nessuna riga da copiare.")
        if not (html_table := self._generate_html_table_with_totals(pages_to_copy)): return
        try:
            self._set_clipboard_html(html_table)
            self.status_text.set(f"Tabella formattata ({len(pages_to_copy)} righe) copiata.")
        except Exception as e:
            self.status_text.set("Errore: impossibile copiare la tabella.")
            messagebox.showerror("Errore Appunti", f"Impossibile copiare l'HTML.\n{e}", parent=self)
            traceback.print_exc()

    def export_to_pdf(self, selection_mode=False):
        pages_to_export = self.get_pages_for_selection(selection_mode)
        if not pages_to_export: return messagebox.showinfo("Informazione", "Nessun dato da esportare.", parent=self)
        self._lock_ui()
        options_dialog = ExportOptionsWindow(self)
        self.wait_window(options_dialog)
        if not (options := options_dialog.result): self._unlock_ui(); return self.status_text.set("Esportazione annullata.")
        file_path = filedialog.asksaveasfilename(parent=self, defaultextension=".pdf", filetypes=[("File PDF", "*.pdf")], title="Salva report PDF")
        if not file_path: self._unlock_ui(); return self.status_text.set("Esportazione PDF annullata.")
        self.status_text.set("Creazione PDF in corso..."); self.update_idletasks()
        threading.Thread(target=self._build_pdf_thread, args=(file_path, options, pages_to_export), daemon=True).start()

    def _build_pdf_thread(self, file_path, options, pages_to_export):
        try:
            grouped_pages = defaultdict(list)
            for page in pages_to_export: grouped_pages[page['file_info']['scan_root']].append(page)
            page_size = landscape(A4) if options['orientation'] == 'landscape' else A4
            doc = SimpleDocTemplate(file_path, pagesize=page_size, topMargin=1.5*cm, bottomMargin=1.5*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
            styles = getSampleStyleSheet()
            filename_style = ParagraphStyle('file_style', parent=styles['Normal'], fontSize=8, alignment=1)
            dims_style = ParagraphStyle('dims_style', parent=styles['Normal'], fontSize=7, textColor=colors.darkgrey, alignment=1)
            trim_dims_style = ParagraphStyle('trim_dims_style', parent=styles['Normal'], fontSize=6, textColor=colors.red, alignment=1)
            folder_header_style = ParagraphStyle('folder_header', parent=styles['h2'], backColor=colors.lightblue, padding=4, textColor=colors.black)
            story, num_columns = [], options['columns']
            col_width = (doc.width / num_columns) - (cm * 0.2 * (num_columns - 1))
            sorted_grouped_pages = sorted(grouped_pages.items(), key=lambda item: item[0])

            for folder, pages in sorted_grouped_pages:
                display_folder = os.path.basename(folder)
                num_files = len({(p['file_info']['path'], p['file_info']['filename']) for p in pages})
                total_pages = len(pages)
                total_sqm = sum(p['file_info']['pages_details'][p['page_num']]['area_sqm'] for p in pages)
                total_trim_sqm = sum(p['file_info']['pages_details'][p['page_num']].get('trim_area_sqm', p['file_info']['pages_details'][p['page_num']]['area_sqm']) for p in pages)
                
                area_text = f"Area: {total_sqm:.2f} m²"
                if abs(total_sqm - total_trim_sqm) > 0.0001:
                    area_text += f" (Al vivo: {total_trim_sqm:.2f} m²)"

                story.extend([Paragraph(display_folder, folder_header_style), Paragraph(f"File: {num_files} | Pagine: {total_pages} | {area_text}", styles['Normal']), Spacer(1, 0.5*cm)])
                grid_data, row = [], []
                for page_data in pages:
                    item_data, page_num = page_data['file_info'], page_data['page_num']
                    page_details = item_data['pages_details'][page_num]
                    full_path = os.path.join(item_data['path'], item_data['filename'])
                    cell_content = []
                    try:
                        if item_data['type'] in ('PDF', 'AI'):
                            with fitz.open(full_path) as doc_pdf:
                                img_data = doc_pdf.load_page(page_num).get_pixmap(dpi=150).tobytes("png")
                                img_report = ReportLabImage(io.BytesIO(img_data), width=col_width*0.9, height=col_width*0.9, kind='proportional')
                        else:
                            with Image.open(full_path) as img:
                                img.thumbnail((400, 400)); img_buffer = io.BytesIO()
                                img.save(img_buffer, format='PNG'); img_buffer.seek(0)
                                img_report = ReportLabImage(img_buffer, width=col_width*0.9, height=col_width*0.9, kind='proportional')
                        cell_content.append(img_report)
                    except Exception as e: print(f"Errore anteprima PDF per {full_path}: {e}")
                    page_info = f" (Pag. {page_num + 1}/{item_data.get('page_count', 1)})" if item_data.get('page_count', 1) > 1 else ""
                    cell_content.append(Paragraph(item_data['filename'] + page_info, filename_style))
                    dpi_info = f"({item_data['dpi_str']})" if item_data.get('dpi_str') else ""
                    cell_content.append(Paragraph(page_details['dimensions_cm'] + f" cm {dpi_info}", dims_style))
                    if 'trim_dimensions_cm' in page_details:
                        cell_content.append(Paragraph(f"Al vivo: {page_details['trim_dimensions_cm']} cm", trim_dims_style))

                    row.append(cell_content)
                    if len(row) == num_columns: grid_data.append(row); row = []
                if row: row.extend([""] * (num_columns - len(row))); grid_data.append(row)
                if grid_data:
                    story.append(Table(grid_data, colWidths=[col_width]*num_columns, style=[('VALIGN',(0,0),(-1,-1),'TOP'), ('ALIGN',(0,0),(-1,-1),'CENTER'), ('BOX',(0,0),(-1,-1),1,colors.lightgrey), ('PADDING',(0,0),(-1,-1),6)]))
                story.append(PageBreak())
            if story and isinstance(story[-1], PageBreak): story.pop()
            doc.build(story)
            self.after(0, self.on_pdf_success, file_path)
        except Exception as e: self.after(0, self.on_pdf_error, e)

    def on_pdf_success(self, file_path):
        self.status_text.set("Report PDF creato con successo.")
        self._unlock_ui(); self.update()
        messagebox.showinfo("Successo", f"Report esportato con successo in:\n{file_path}", parent=self)

    def on_pdf_error(self, e):
        traceback.print_exc(); self.status_text.set("Errore creazione PDF.")
        self._unlock_ui(); self.update()
        messagebox.showerror("Errore", f"Impossibile creare il PDF.\n{e}", parent=self)

def create_tab(tab_view):
    tab_name = "Liste e anteprime"
    tab = tab_view.add(tab_name)
    app_instance = FileScannerApp(master=tab)
    return tab_name, app_instance
