# app_controllo_immagini.py - v9.0
import customtkinter as ctk
import os
from tkinter import filedialog, Canvas, Toplevel
from PIL import Image, ExifTags, ImageOps, ImageTk
import math

# Disabilita il limite di dimensione per le immagini grandi
Image.MAX_IMAGE_PIXELS = None

class ImageCheckerApp(ctk.CTkFrame):
    """
    Applicazione per visualizzare, analizzare e trasformare immagini.
    """
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        self.loaded_images = []
        self.thumbnail_buttons = []
        self.active_image_index = -1
        self.original_image_obj = None 
        self.modified_image_obj = None 
        self.photo_image = None 
        self.image_width_px = None
        self.image_height_px = None
        self._is_updating_dimensions = False

        # --- Variabili per il ritaglio interattivo ---
        self.crop_rect_coords = {} 
        self.crop_rect_id = None
        self.crop_handles = {}
        self.active_handle = None
        self.drag_mode = None 
        self.drag_start_pos = (0, 0)
        self.image_on_canvas_id = None
        self.canvas_scale_factor = 1.0
        self.canvas_offset = (0, 0)

        # --- NUOVO: Variabili per la Lente d'Ingrandimento ---
        self.magnifier_window = None
        self.magnifier_canvas = None
        self.magnifier_photo = None
        self.magnifier_size = 100  # Dimensione della lente in pixel
        self.magnifier_zoom = 5    # Fattore di ingrandimento

        # --- Layout Principale ---
        self.grid_columnconfigure(0, weight=0, minsize=180)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Frame Lista Miniature (Sinistra) ---
        self.thumbnail_list_frame = ctk.CTkScrollableFrame(self, label_text="Immagini")
        self.thumbnail_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- Frame Contenuto Principale (Destra) ---
        main_content_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_content_frame.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        main_content_frame.grid_columnconfigure(0, weight=1)
        main_content_frame.grid_rowconfigure(0, weight=0)
        main_content_frame.grid_rowconfigure(1, weight=1)

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
        bottom_content_frame.grid_columnconfigure(0, weight=1, minsize=350)
        bottom_content_frame.grid_columnconfigure(1, weight=4)
        bottom_content_frame.grid_rowconfigure(0, weight=1)

        # --- Frame Colonna Sinistra (Dettagli + Trasformazioni) ---
        left_column_frame = ctk.CTkFrame(bottom_content_frame, fg_color="transparent")
        left_column_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        left_column_frame.grid_rowconfigure(0, weight=0)
        left_column_frame.grid_rowconfigure(1, weight=0)
        left_column_frame.grid_columnconfigure(0, weight=1)

        # --- Frame Dettagli (in alto a sx) ---
        details_frame = ctk.CTkFrame(left_column_frame)
        details_frame.grid(row=0, column=0, sticky="new")
        details_frame.grid_columnconfigure(1, weight=1)

        # --- Frame Anteprima con Canvas ---
        preview_frame = ctk.CTkFrame(bottom_content_frame)
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_canvas = Canvas(preview_frame, bg="#2b2b2b", highlightthickness=0, relief='ridge', cursor="cross")
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        
        self.preview_canvas.bind("<ButtonPress-1>", self._on_canvas_press)
        self.preview_canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self.preview_canvas.bind("<Motion>", self._on_canvas_motion)

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

        apply_dims_btn = ctk.CTkButton(details_frame, text="Applica Dimensioni", command=self._apply_dimensions_and_resample)
        apply_dims_btn.grid(row=row_index, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_index += 1

        self.distance_var = ctk.StringVar(value="-")
        distance_label = ctk.CTkLabel(details_frame, text="Distanza visione:", anchor="e", font=ctk.CTkFont(weight="bold"))
        distance_label.grid(row=row_index, column=0, padx=10, pady=5, sticky="e")
        distance_value_label = ctk.CTkLabel(details_frame, textvariable=self.distance_var, anchor="w")
        distance_value_label.grid(row=row_index, column=1, padx=10, pady=5, sticky="w")

        # --- Frame Trasformazioni e Salvataggio ---
        transform_frame = ctk.CTkFrame(left_column_frame)
        transform_frame.grid(row=1, column=0, sticky="new", pady=(10, 0))
        transform_frame.grid_columnconfigure(0, weight=1)

        transform_buttons_frame = ctk.CTkFrame(transform_frame, fg_color="transparent")
        transform_buttons_frame.grid(row=1, column=0, sticky="ew")
        transform_buttons_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        rotate_left_btn = ctk.CTkButton(transform_buttons_frame, text="↶", command=lambda: self._rotate_image(90), width=40)
        rotate_left_btn.grid(row=0, column=0, padx=(5,2), pady=5, sticky="ew")
        
        flip_hor_btn = ctk.CTkButton(transform_buttons_frame, text="↔", command=lambda: self._flip_image("horizontal"), width=40)
        flip_hor_btn.grid(row=0, column=1, padx=2, pady=5, sticky="ew")

        flip_ver_btn = ctk.CTkButton(transform_buttons_frame, text="↕", command=lambda: self._flip_image("vertical"), width=40)
        flip_ver_btn.grid(row=0, column=2, padx=2, pady=5, sticky="ew")

        rotate_right_btn = ctk.CTkButton(transform_buttons_frame, text="↷", command=lambda: self._rotate_image(-90), width=40)
        rotate_right_btn.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew")

        # --- Sezione Ritaglio ---
        crop_input_frame = ctk.CTkFrame(transform_frame, fg_color="transparent")
        crop_input_frame.grid(row=2, column=0, sticky="ew", padx=5)
        crop_input_frame.grid_columnconfigure((1, 3), weight=1)

        self.crop_width_cm_var = ctk.StringVar()
        self.crop_height_cm_var = ctk.StringVar()

        crop_w_label = ctk.CTkLabel(crop_input_frame, text="L (cm):")
        crop_w_label.grid(row=0, column=0, padx=(0, 2))
        self.crop_width_cm_entry = ctk.CTkEntry(crop_input_frame, textvariable=self.crop_width_cm_var)
        self.crop_width_cm_entry.grid(row=0, column=1, sticky="ew")
        
        crop_h_label = ctk.CTkLabel(crop_input_frame, text="A (cm):")
        crop_h_label.grid(row=0, column=2, padx=(8, 2))
        self.crop_height_cm_entry = ctk.CTkEntry(crop_input_frame, textvariable=self.crop_height_cm_var)
        self.crop_height_cm_entry.grid(row=0, column=3, sticky="ew")

        set_crop_btn = ctk.CTkButton(crop_input_frame, text="Imposta", command=self._set_crop_from_entries, width=60)
        set_crop_btn.grid(row=0, column=4, padx=(5,0))

        crop_buttons_frame = ctk.CTkFrame(transform_frame, fg_color="transparent")
        crop_buttons_frame.grid(row=3, column=0, columnspan=2, sticky="ew")
        crop_buttons_frame.grid_columnconfigure(0, weight=1)
        crop_buttons_frame.grid_columnconfigure(1, weight=1)

        crop_btn = ctk.CTkButton(crop_buttons_frame, text="Applica Ritaglio", command=self._crop_image)
        crop_btn.grid(row=0, column=0, padx=(5,2), pady=5, sticky="ew")
        
        reset_crop_btn = ctk.CTkButton(crop_buttons_frame, text="Reset Ritaglio", command=self._reset_crop_to_full_image, fg_color="gray50", hover_color="gray30")
        reset_crop_btn.grid(row=0, column=1, padx=(2,5), pady=5, sticky="ew")

        save_btn = ctk.CTkButton(transform_frame, text="Salva con nome...", command=self._save_image_as, fg_color="#28a745", hover_color="#218838")
        save_btn.grid(row=4, column=0, columnspan=2, padx=5, pady=(10, 5), sticky="ew")

        self._create_magnifier()


    def handle_drop(self, event):
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
        try:
            with Image.open(image_path) as img:
                img = ImageOps.exif_transpose(img.copy())
                width_px, height_px = img.size
                gcd = math.gcd(width_px, height_px)
                file_size = os.path.getsize(image_path)
                if file_size < 1024: size_str = f"{file_size} Bytes"
                elif file_size < 1024**2: size_str = f"{file_size/1024:.2f} KB"
                else: size_str = f"{file_size/1024**2:.2f} MB"
                return {
                    "path": image_path, "image_obj": img, "width_px": width_px,
                    "height_px": height_px, "Nome file:": os.path.basename(image_path),
                    "Percorso:": os.path.dirname(image_path), "Dimensioni:": f"{width_px} x {height_px} px",
                    "Modo colore:": img.mode, "Dimensione file:": size_str,
                    "Rapporto:": f"{width_px//gcd}:{height_px//gcd}",
                    "Risoluzione:": img.info.get('dpi', (96, 96))[0]
                }
        except Exception as e:
            print(f"Errore nell'analisi dell'immagine {image_path}: {e}")
            return None

    def _populate_thumbnail_list(self):
        for widget in self.thumbnail_list_frame.winfo_children():
            widget.destroy()
        self.thumbnail_buttons.clear()
        for index, img_data in enumerate(self.loaded_images):
            img_copy = img_data["image_obj"].copy()
            img_copy.thumbnail((150, 150), Image.Resampling.LANCZOS)
            thumb_img = ctk.CTkImage(light_image=img_copy, dark_image=img_copy, size=img_copy.size)
            btn = ctk.CTkButton(self.thumbnail_list_frame, image=thumb_img, text=img_data["Nome file:"], 
                                compound="top", font=ctk.CTkFont(size=11),
                                command=lambda idx=index: self._display_image_details(idx))
            btn.pack(padx=5, pady=5, fill="x")
            self.thumbnail_buttons.append(btn)
        if self.loaded_images: self._display_image_details(0)
        else: self._clear_details()

    def _display_image_details(self, index):
        self.active_image_index = index
        image_info = self.loaded_images[index]
        self.original_image_obj = image_info["image_obj"].copy()
        self.modified_image_obj = image_info["image_obj"].copy()
        for i, btn in enumerate(self.thumbnail_buttons):
            btn.configure(fg_color="transparent" if i != index else ctk.ThemeManager.theme["CTkButton"]["fg_color"])
        self._update_details_and_preview()

    def _update_details_and_preview(self):
        if self.modified_image_obj is None: return
        self.image_width_px, self.image_height_px = self.modified_image_obj.size
        original_info = self.loaded_images[self.active_image_index]
        self.info_vars["Nome file:"].set(original_info["Nome file:"])
        self.info_vars["Percorso:"].set(original_info["Percorso:"])
        self.info_vars["Modo colore:"].set(self.modified_image_obj.mode)
        self.info_vars["Dimensione file:"].set(original_info["Dimensione file:"])
        gcd = math.gcd(self.image_width_px, self.image_height_px) if self.image_width_px > 0 and self.image_height_px > 0 else 1
        self.info_vars["Dimensioni:"].set(f"{self.image_width_px} x {self.image_height_px} px")
        self.info_vars["Rapporto:"].set(f"{self.image_width_px//gcd}:{self.image_height_px//gcd}")
        
        self._is_updating_dimensions = True
        initial_dpi = self.modified_image_obj.info.get('dpi', (96, 96))[0]
        self.dpi_var.set(str(int(initial_dpi)))
        self._update_cm_from_pixels(self.image_width_px, self.image_height_px, int(initial_dpi))
        self._is_updating_dimensions = False
        
        self.after(50, self._update_preview_canvas)
        self._update_distance_from_dpi()

    def _update_preview_canvas(self):
        if not self.modified_image_obj: return
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        if canvas_width < 20 or canvas_height < 20: return
        img_w, img_h = self.modified_image_obj.size
        self.canvas_scale_factor = min(canvas_width / img_w, canvas_height / img_h) if img_w > 0 and img_h > 0 else 1
        thumb_w = int(img_w * self.canvas_scale_factor)
        thumb_h = int(img_h * self.canvas_scale_factor)
        thumb = self.modified_image_obj.resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
        self.photo_image = ImageTk.PhotoImage(thumb)
        self.preview_canvas.delete("all")
        self.canvas_offset = (int((canvas_width - thumb_w) / 2), int((canvas_height - thumb_h) / 2))
        self.image_on_canvas_id = self.preview_canvas.create_image(
            self.canvas_offset[0], self.canvas_offset[1], anchor="nw", image=self.photo_image
        )
        self._reset_crop()

    def _clear_details(self):
        self._is_updating_dimensions = True
        for var in self.info_vars.values(): var.set("-")
        self.dpi_var.set("")
        self.width_cm_var.set("")
        self.height_cm_var.set("")
        self.distance_var.set("-")
        self.crop_width_cm_var.set("")
        self.crop_height_cm_var.set("")
        self._is_updating_dimensions = False
        self.preview_canvas.delete("all")
        self.image_width_px = None
        self.modified_image_obj = None
        self.original_image_obj = None
    
    # --- Funzioni di calcolo e aggiornamento UI ---

    def _update_cm_from_pixels(self, px_w, px_h, dpi):
        if dpi > 0:
            self.width_cm_var.set(f"{(px_w / dpi) * 2.54:.2f}")
            self.height_cm_var.set(f"{(px_h / dpi) * 2.54:.2f}")

    def _update_dpi_from_cm(self, px_w, cm_w):
        if cm_w > 0:
            self.dpi_var.set(str(int((px_w / cm_w) * 2.54)))

    def _update_distance_from_dpi(self):
        try:
            dpi = int(self.dpi_var.get())
            if dpi > 0:
                distance_m = 80 / dpi
                self.distance_var.set(f"~ {distance_m:.2f} m")
        except (ValueError, ZeroDivisionError):
            self.distance_var.set("-")

    def _update_from_width(self, *args):
        if self._is_updating_dimensions or not self.image_width_px: return
        self._is_updating_dimensions = True
        try:
            w_cm = float(self.width_cm_var.get().replace(",", "."))
            if w_cm > 0 and self.image_width_px > 0:
                aspect_ratio = self.image_height_px / self.image_width_px
                self.height_cm_var.set(f"{w_cm * aspect_ratio:.2f}")
                self._update_dpi_from_cm(self.image_width_px, w_cm)
                self._update_distance_from_dpi()
        except (ValueError, ZeroDivisionError): pass
        finally: self._is_updating_dimensions = False

    def _update_from_height(self, *args):
        if self._is_updating_dimensions or not self.image_height_px: return
        self._is_updating_dimensions = True
        try:
            h_cm = float(self.height_cm_var.get().replace(",", "."))
            if h_cm > 0 and self.image_height_px > 0:
                aspect_ratio = self.image_width_px / self.image_height_px
                self.width_cm_var.set(f"{h_cm * aspect_ratio:.2f}")
                self._update_dpi_from_cm(self.image_height_px, h_cm)
                self._update_distance_from_dpi()
        except (ValueError, ZeroDivisionError): pass
        finally: self._is_updating_dimensions = False

    def _update_from_dpi(self, *args):
        if self._is_updating_dimensions: return
        self._is_updating_dimensions = True
        try:
            dpi = int(self.dpi_var.get())
            self._update_cm_from_pixels(self.image_width_px, self.image_height_px, dpi)
            self._update_distance_from_dpi()
        except ValueError: pass
        finally: self._is_updating_dimensions = False

    # --- Funzioni di trasformazione effettiva ---

    def _apply_dimensions_and_resample(self):
        if not self.modified_image_obj: return
        try:
            w_cm = float(self.width_cm_var.get().replace(",", "."))
            h_cm = float(self.height_cm_var.get().replace(",", "."))
            dpi = int(self.dpi_var.get())

            if w_cm <= 0 or h_cm <= 0 or dpi <= 0:
                print("Valori per il ridimensionamento non validi.")
                return

            new_width_px = int(round((w_cm / 2.54) * dpi))
            new_height_px = int(round((h_cm / 2.54) * dpi))

            self.modified_image_obj = self.modified_image_obj.resize(
                (new_width_px, new_height_px), Image.Resampling.LANCZOS
            )
            self.modified_image_obj.info['dpi'] = (dpi, dpi)
            
            self._update_details_and_preview()

        except ValueError:
            print("Errore: Inserire valori numerici validi per ridimensionare.")


    def _flip_image(self, direction):
        if self.modified_image_obj:
            if direction == "horizontal": self.modified_image_obj = self.modified_image_obj.transpose(Image.FLIP_LEFT_RIGHT)
            elif direction == "vertical": self.modified_image_obj = self.modified_image_obj.transpose(Image.FLIP_TOP_BOTTOM)
            self._update_details_and_preview()

    def _rotate_image(self, angle):
        if self.modified_image_obj:
            self.modified_image_obj = self.modified_image_obj.rotate(angle, expand=True)
            self._update_details_and_preview()

    def _save_image_as(self):
        if not self.modified_image_obj: return
        original_info = self.loaded_images[self.active_image_index]
        original_path = original_info["path"]
        directory = os.path.dirname(original_path)
        filename, extension = os.path.splitext(os.path.basename(original_path))
        suggested_filename = f"{filename}_modificato{extension}"
        file_path = filedialog.asksaveasfilename(
            initialdir=directory, initialfile=suggested_filename, defaultextension=extension,
            filetypes=[("JPEG", "*.jpg;*.jpeg"), ("PNG", "*.png"), ("BMP", "*.bmp"), ("Tutti i file", "*.*")]
        )
        if file_path:
            try:
                dpi_val = int(self.dpi_var.get())
                self.modified_image_obj.save(file_path, dpi=(dpi_val, dpi_val))
                print(f"Immagine salvata con successo in: {file_path}")
            except Exception as e: print(f"Errore durante il salvataggio dell'immagine: {e}")

    # --- Funzioni per il Ritaglio Interattivo e Lente ---

    def _create_magnifier(self):
        self.magnifier_window = Toplevel(self)
        self.magnifier_window.overrideredirect(True)
        self.magnifier_window.wm_attributes("-topmost", True)
        self.magnifier_canvas = Canvas(self.magnifier_window, width=self.magnifier_size, height=self.magnifier_size, highlightthickness=0)
        self.magnifier_canvas.pack()
        self.magnifier_window.withdraw()

    def _show_magnifier(self, event):
        if self.magnifier_window:
            self.magnifier_window.deiconify()
            self._update_magnifier(event)

    def _hide_magnifier(self):
        if self.magnifier_window:
            self.magnifier_window.withdraw()

    def _update_magnifier(self, event):
        if not self.modified_image_obj:
            self._hide_magnifier()
            return

        # Posiziona la lente vicino al cursore
        self.magnifier_window.geometry(f"+{event.x_root + 20}+{event.y_root + 20}")

        # Converte le coordinate del canvas in coordinate dell'immagine originale
        img_x = int((event.x - self.canvas_offset[0]) / self.canvas_scale_factor)
        img_y = int((event.y - self.canvas_offset[1]) / self.canvas_scale_factor)

        # Ritaglia una piccola area dall'immagine originale
        box_size = int(self.magnifier_size / self.magnifier_zoom)
        box = (
            img_x - box_size // 2,
            img_y - box_size // 2,
            img_x + box_size // 2,
            img_y + box_size // 2,
        )
        
        try:
            region = self.modified_image_obj.crop(box)
            # Ingrandisce la regione
            zoomed_region = region.resize((self.magnifier_size, self.magnifier_size), Image.Resampling.NEAREST)
            self.magnifier_photo = ImageTk.PhotoImage(zoomed_region)
            self.magnifier_canvas.create_image(0, 0, anchor="nw", image=self.magnifier_photo)
            
            # Disegna il mirino
            center = self.magnifier_size / 2
            self.magnifier_canvas.create_line(center, 0, center, self.magnifier_size, fill="red")
            self.magnifier_canvas.create_line(0, center, self.magnifier_size, center, fill="red")
        except Exception:
            # Se il cursore è fuori dall'immagine, non fare nulla
            pass

    def _reset_crop(self):
        self.preview_canvas.delete(self.crop_rect_id)
        for handle in self.crop_handles.values():
            self.preview_canvas.delete(handle)
        self.crop_rect_id = None
        self.crop_handles = {}
        self.crop_rect_coords = {}
        self.drag_mode = None
        self.active_handle = None
        self.crop_width_cm_var.set("")
        self.crop_height_cm_var.set("")

    def _reset_crop_to_full_image(self):
        if not self.modified_image_obj or not self.photo_image: return
        
        x1 = self.canvas_offset[0]
        y1 = self.canvas_offset[1]
        x2 = x1 + self.photo_image.width()
        y2 = y1 + self.photo_image.height()

        self.crop_rect_coords = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2}
        self._update_crop_display()

    def _update_crop_display(self):
        if not self.crop_rect_coords: return
        
        c = self.crop_rect_coords
        if self.crop_rect_id:
            self.preview_canvas.coords(self.crop_rect_id, c['x1'], c['y1'], c['x2'], c['y2'])
        else:
            self.crop_rect_id = self.preview_canvas.create_rectangle(
                c['x1'], c['y1'], c['x2'], c['y2'],
                outline="red", width=2, tags="crop_rect")

        handle_size = 8
        handles_pos = {
            'n': ( (c['x1']+c['x2'])/2, c['y1'] ), 's': ( (c['x1']+c['x2'])/2, c['y2'] ),
            'w': ( c['x1'], (c['y1']+c['y2'])/2 ), 'e': ( c['x2'], (c['y1']+c['y2'])/2 )
        }
        for handle, pos in handles_pos.items():
            x, y = pos
            if handle in self.crop_handles:
                self.preview_canvas.coords(self.crop_handles[handle], x-handle_size/2, y-handle_size/2, x+handle_size/2, y+handle_size/2)
            else:
                self.crop_handles[handle] = self.preview_canvas.create_rectangle(
                    x-handle_size/2, y-handle_size/2, x+handle_size/2, y+handle_size/2,
                    fill="red", outline="white", tags=(f"handle_{handle}", "handle")
                )
        
        try:
            _, _, w_orig, h_orig = self._get_original_crop_coords()
            dpi = int(self.dpi_var.get()) if self.dpi_var.get() else 96
            if dpi <= 0: dpi = 96
            
            w_cm = (w_orig / dpi) * 2.54
            h_cm = (h_orig / dpi) * 2.54
            self.crop_width_cm_var.set(f"{w_cm:.2f}")
            self.crop_height_cm_var.set(f"{h_cm:.2f}")

        except Exception:
             self.crop_width_cm_var.set("")
             self.crop_height_cm_var.set("")

    def _set_crop_from_entries(self):
        if not self.modified_image_obj: return
        try:
            w_cm = float(self.crop_width_cm_var.get().replace(",", "."))
            h_cm = float(self.crop_height_cm_var.get().replace(",", "."))
            dpi = int(self.dpi_var.get())

            if w_cm <= 0 or h_cm <= 0 or dpi <= 0: return

            crop_w_px = (w_cm / 2.54) * dpi
            crop_h_px = (h_cm / 2.54) * dpi

            crop_w_canvas = crop_w_px * self.canvas_scale_factor
            crop_h_canvas = crop_h_px * self.canvas_scale_factor

            center_x = self.canvas_offset[0] + self.photo_image.width() / 2
            center_y = self.canvas_offset[1] + self.photo_image.height() / 2
            
            x1 = center_x - crop_w_canvas / 2
            y1 = center_y - crop_h_canvas / 2
            x2 = center_x + crop_w_canvas / 2
            y2 = center_y + crop_h_canvas / 2

            self.crop_rect_coords = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2}
            self._update_crop_display()

        except (ValueError, ZeroDivisionError):
            print("Valori per il ritaglio non validi.")


    def _on_canvas_motion(self, event):
        if self.drag_mode: return
        
        active_items = self.preview_canvas.find_withtag("current")
        tags = self.preview_canvas.gettags(active_items[0]) if active_items else []
        
        new_cursor = "cross"
        if "handle" in tags:
            if "handle_n" in tags or "handle_s" in tags: new_cursor = "sb_v_double_arrow"
            elif "handle_w" in tags or "handle_e" in tags: new_cursor = "sb_h_double_arrow"
        elif "crop_rect" in tags:
            new_cursor = "fleur"
        
        if self.preview_canvas.cget('cursor') != new_cursor:
            self.preview_canvas.config(cursor=new_cursor)

    def _on_canvas_press(self, event):
        self.drag_start_pos = (event.x, event.y)
        active_items = self.preview_canvas.find_withtag("current")
        
        if active_items:
            tags = self.preview_canvas.gettags(active_items[0])
            if "handle" in tags:
                self.drag_mode = "resize"
                self.active_handle = tags[0].split('_')[1]
                self._show_magnifier(event)
                return
            elif "crop_rect" in tags:
                self.drag_mode = "move"
                return

        self._reset_crop()
        self.drag_mode = "draw"
        self.crop_rect_coords = {'x1': event.x, 'y1': event.y, 'x2': event.x, 'y2': event.y}
        self._update_crop_display()

    def _on_canvas_drag(self, event):
        if not self.drag_mode: return
        
        if self.drag_mode == "resize":
            self._update_magnifier(event)
        
        dx = event.x - self.drag_start_pos[0]
        dy = event.y - self.drag_start_pos[1]
        c = self.crop_rect_coords

        if self.drag_mode == "draw":
            c['x2'] = event.x
            c['y2'] = event.y
        elif self.drag_mode == "resize":
            if self.active_handle == 'n': c['y1'] += dy
            elif self.active_handle == 's': c['y2'] += dy
            elif self.active_handle == 'w': c['x1'] += dx
            elif self.active_handle == 'e': c['x2'] += dx
            self.drag_start_pos = (event.x, event.y)
        elif self.drag_mode == "move":
            c['x1'] += dx
            c['x2'] += dx
            c['y1'] += dy
            c['y2'] += dy
            self.drag_start_pos = (event.x, event.y)
        
        self._update_crop_display()

    def _on_canvas_release(self, event):
        self._hide_magnifier()
        c = self.crop_rect_coords
        if c and c['x1'] > c['x2']: c['x1'], c['x2'] = c['x2'], c['x1']
        if c and c['y1'] > c['y2']: c['y1'], c['y2'] = c['y2'], c['y1']
        
        self.drag_mode = None
        self.active_handle = None
        if self.crop_rect_id:
            self._update_crop_display()

    def _get_original_crop_coords(self):
        if not self.crop_rect_coords: return None
        c = self.crop_rect_coords
        
        img_x1 = int((c['x1'] - self.canvas_offset[0]) / self.canvas_scale_factor)
        img_y1 = int((c['y1'] - self.canvas_offset[1]) / self.canvas_scale_factor)
        img_x2 = int((c['x2'] - self.canvas_offset[0]) / self.canvas_scale_factor)
        img_y2 = int((c['y2'] - self.canvas_offset[1]) / self.canvas_scale_factor)
        
        img_w, img_h = self.modified_image_obj.size
        img_x1 = max(0, img_x1)
        img_y1 = max(0, img_y1)
        img_x2 = min(img_w, img_x2)
        img_y2 = min(img_h, img_y2)
        
        width_orig = img_x2 - img_x1
        height_orig = img_y2 - img_y1

        return (img_x1, img_y1, width_orig, height_orig)

    def _crop_image(self):
        if self.modified_image_obj and self.crop_rect_id:
            coords = self._get_original_crop_coords()
            if coords and coords[2] > 0 and coords[3] > 0:
                x, y, w, h = coords
                self.modified_image_obj = self.modified_image_obj.crop((x, y, x + w, y + h))
                self._update_details_and_preview()
            else:
                print("Ritaglio non valido: dimensioni nulle.")
                self._reset_crop()

# --- Funzione di caricamento richiesta da winfile.py ---
def create_tab(tab_view):
    tab_name = "Controllo immagini"
    tab = tab_view.add(tab_name)
    app_instance = ImageCheckerApp(master=tab)
    return tab_name, app_instance
