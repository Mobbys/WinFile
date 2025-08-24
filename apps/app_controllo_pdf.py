# apps/app_controllo_pdf.py - v8.0 (Solo Ritaglio)
import customtkinter as ctk
import os
import fitz  # PyMuPDF
from tkinter import messagebox, filedialog
import io
from PIL import Image
import traceback

# Prova a importare la libreria per il Drag & Drop
try:
    from tkinterdnd2 import DND_FILES
    DND_SUPPORT = True
except ImportError:
    DND_SUPPORT = False

# Disabilita il limite di dimensione per le immagini grandi
Image.MAX_IMAGE_PIXELS = None

class PDFCheckerApp(ctk.CTkFrame):
    """
    Applicazione per analizzare file PDF, visualizzare le pagine
    e ritagliare le pagine in base al TrimBox.
    """
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        self.doc = None
        self.doc_path = None
        self.modified_doc = None
        self.active_page_index = -1
        self.ctk_image = None
        self.page_thumbnail_buttons = []
        self.zoom_level = 1.0 # Livello di zoom iniziale
        self.pan_start_x = 0
        self.pan_start_y = 0

        self.grid_columnconfigure(0, weight=0, minsize=180)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.thumbnail_list_frame = ctk.CTkScrollableFrame(self, label_text="Pagine")
        self.thumbnail_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_content_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        main_content_frame.grid_columnconfigure(0, weight=1)
        main_content_frame.grid_rowconfigure(0, weight=0)
        main_content_frame.grid_rowconfigure(1, weight=1)

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

        bottom_content_frame = ctk.CTkFrame(main_content_frame, fg_color="transparent")
        bottom_content_frame.grid(row=1, column=0, pady=(10, 0), sticky="nsew")
        bottom_content_frame.grid_columnconfigure(0, weight=2, minsize=350)
        bottom_content_frame.grid_columnconfigure(1, weight=4)
        bottom_content_frame.grid_rowconfigure(0, weight=1) # Riga per l'anteprima
        bottom_content_frame.grid_rowconfigure(1, weight=0) # Riga per i controlli zoom

        details_actions_frame = ctk.CTkFrame(bottom_content_frame)
        details_actions_frame.grid(row=0, column=0, rowspan=2, padx=(0, 10), sticky="nsew")
        details_actions_frame.grid_columnconfigure(0, weight=1)
        
        self.preview_frame = ctk.CTkScrollableFrame(bottom_content_frame)
        self.preview_frame.grid(row=0, column=1, sticky="nsew")

        self.image_label = ctk.CTkLabel(self.preview_frame, text="", text_color="gray")
        self.image_label.pack(expand=True)
        
        if DND_SUPPORT:
            drop_target = self.preview_frame if not self.image_label.winfo_viewable() else self.image_label
            drop_target.drop_target_register(DND_FILES)
            drop_target.dnd_bind('<<Drop>>', self.handle_drop)
            self.image_label.configure(text="Trascina un file PDF qui")
        else:
            self.image_label.configure(text="Funzionalità di trascinamento disabilitata.")

        self.image_label.bind("<ButtonPress-1>", self._on_pan_start)
        self.image_label.bind("<B1-Motion>", self._on_pan_move)
        self.image_label.bind("<ButtonRelease-1>", self._on_pan_end)
        self.image_label.configure(cursor="hand2")


        zoom_frame = ctk.CTkFrame(bottom_content_frame, fg_color="transparent")
        zoom_frame.grid(row=1, column=1, pady=(5,0), sticky="ew")
        zoom_frame.grid_columnconfigure((0, 2), weight=1)
        zoom_frame.grid_columnconfigure(1, weight=0)

        zoom_out_button = ctk.CTkButton(zoom_frame, text="-", width=40, command=self.zoom_out)
        zoom_out_button.grid(row=0, column=0, sticky="e")

        self.zoom_label = ctk.CTkLabel(zoom_frame, text="100%", width=60)
        self.zoom_label.grid(row=0, column=1, padx=10)

        zoom_in_button = ctk.CTkButton(zoom_frame, text="+", width=40, command=self.zoom_in)
        zoom_in_button.grid(row=0, column=2, sticky="w")

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
        actions_frame.pack(fill="both", expand=True, pady=10)
        
        self.trim_button = ctk.CTkButton(actions_frame, text="Rifila a TrimBox", command=self.crop_to_trimbox, state="disabled")
        self.trim_button.pack(pady=5)
        self.save_button = ctk.CTkButton(actions_frame, text="Salva PDF Modificato...", command=self.save_modified_pdf, state="disabled")
        self.save_button.pack(pady=5)

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
            path_string = event.data.strip()
            if path_string.startswith('{') and path_string.endswith('}'):
                path_string = path_string[1:-1]

            possible_paths = self.tk.splitlist(path_string)
            
            file_path = ""
            if possible_paths and os.path.exists(possible_paths[0]):
                file_path = possible_paths[0]
            elif os.path.exists(path_string):
                file_path = path_string
            else:
                messagebox.showerror("Errore nel percorso", f"Impossibile trovare il file:\n{event.data}", parent=self)
                return

            if file_path.lower().endswith('.pdf'):
                self._process_pdf(file_path)
            else:
                messagebox.showwarning("File non supportato", f"Il file '{os.path.basename(file_path)}' non è un PDF.", parent=self)
        except Exception as e:
            messagebox.showerror("Errore Drag & Drop", f"Si è verificato un errore imprevisto.\n\nDettagli: {e}", parent=self)
            self._clear_all()

    def _process_pdf(self, file_path):
        try:
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
        self.page_thumbnail_buttons.clear()
        current_doc = self.modified_doc if self.modified_doc else self.doc
        if not current_doc: return
        for page_num in range(len(current_doc)):
            thumb_img = self._get_page_thumbnail(current_doc[page_num])
            btn = ctk.CTkButton(self.thumbnail_list_frame, image=thumb_img, text=f"Pag. {page_num + 1}", compound="top", font=ctk.CTkFont(size=11), command=lambda idx=page_num: self._display_page_details(idx))
            btn.pack(padx=5, pady=5, fill="x")
            self.page_thumbnail_buttons.append(btn)
        if len(current_doc) > 0: self._display_page_details(0)
        else: self._clear_details()

    def _get_page_thumbnail(self, page, size=(150, 150)):
        pix = page.get_pixmap(dpi=72)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail(size, Image.Resampling.LANCZOS)
        return ctk.CTkImage(light_image=img, dark_image=img, size=img.size)

    def _display_page_details(self, page_index):
        try:
            self.active_page_index = page_index
            current_doc = self.modified_doc if self.modified_doc else self.doc
            page = current_doc[page_index]
            for i, btn in enumerate(self.page_thumbnail_buttons):
                btn.configure(fg_color="transparent" if i != page_index else ctk.ThemeManager.theme["CTkButton"]["fg_color"])
            
            media_box = page.rect
            trim_box = page.trimbox

            self.info_vars["MediaBox (cm):"].set(f"{(media_box.width / 72) * 2.54:.2f} x {(media_box.height / 72) * 2.54:.2f} cm")
            
            if trim_box and trim_box.is_valid:
                 self.info_vars["TrimBox (cm):"].set(f"{(trim_box.width / 72) * 2.54:.2f} x {(trim_box.height / 72) * 2.54:.2f} cm")
            else:
                self.info_vars["TrimBox (cm):"].set("- (non definito)")

            self._update_preview_image()
            
            self.trim_button.configure(state="normal" if page.trimbox != page.mediabox and page.trimbox.is_valid else "disabled")
        except Exception as e:
            messagebox.showerror("Errore Visualizzazione", f"Impossibile mostrare dettagli pagina.\n\nDettagli: {e}", parent=self)

    def zoom_in(self):
        self.zoom_level *= 1.25
        self._update_preview_image()

    def zoom_out(self):
        self.zoom_level /= 1.25
        if self.zoom_level < 0.1: # Limite minimo
            self.zoom_level = 0.1
        self._update_preview_image()

    def _update_preview_image(self):
        if self.active_page_index == -1 or not self.doc:
            return

        current_doc = self.modified_doc if self.modified_doc else self.doc
        page = current_doc[self.active_page_index]
        
        dpi = int(150 * self.zoom_level)
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        self.ctk_image = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
        self.image_label.configure(image=self.ctk_image, text="")
        self.zoom_label.configure(text=f"{self.zoom_level*100:.0f}%")

    def crop_to_trimbox(self):
        if self.active_page_index == -1: return
        try:
            current_doc_to_modify = fitz.open(self.doc_path)
            page = current_doc_to_modify[self.active_page_index]
            if page.trimbox and page.trimbox.is_valid and page.trimbox != page.mediabox:
                page.set_cropbox(page.trimbox)
                self.modified_doc = fitz.open("pdf", current_doc_to_modify.write())
                self._populate_thumbnail_list() 
                self._display_page_details(self.active_page_index)
                self.save_button.configure(state="normal")
                messagebox.showinfo("Rifilatura Applicata", "Salva il PDF per rendere le modifiche permanenti.", parent=self)
            else:
                messagebox.showinfo("Nessuna Azione", "TrimBox coincide già con MediaBox o non è valido.", parent=self)
            current_doc_to_modify.close()
        except Exception as e:
            messagebox.showerror("Errore Rifilatura", f"Impossibile applicare TrimBox.\n\nDettagli: {e}", parent=self)

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

    def _clear_all(self):
        self.doc = None; self.doc_path = None; self.modified_doc = None; self.active_page_index = -1
        for widget in self.thumbnail_list_frame.winfo_children(): widget.destroy()
        self.page_thumbnail_buttons.clear()
        self._clear_details()

    def _clear_details(self):
        for key in ["Nome file:", "Percorso:", "MediaBox (cm):", "TrimBox (cm):"]:
            if key in self.info_vars: self.info_vars[key].set("-")
        self.image_label.configure(image=None, text="Trascina un file PDF qui")
        self.trim_button.configure(state="disabled")
        self.save_button.configure(state="disabled")

def create_tab(tab_view):
    tab_name = "Controllo PDF"
    tab = tab_view.add(tab_name)
    app_instance = PDFCheckerApp(master=tab)
    return tab_name, app_instance
