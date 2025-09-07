# app_simulazione_quote.py
# Versione: 2.1 (7-Set-2025)

# --- Dipendenze Richieste ---
# Per funzionare, questa app richiede le seguenti librerie.
# Puoi installarle con pip:
# pip install customtkinter pillow numpy opencv-python pymupdf ezdxf matplotlib tkinterdnd2

import customtkinter as ctk
from tkinter import Canvas, filedialog, simpledialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import json
import numpy as np
import cv2  # OpenCV per la trasformazione della prospettiva
import io
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- Import per nuovi formati ---
try:
    import fitz  # PyMuPDF per i PDF
except ImportError:
    fitz = None
try:
    import ezdxf
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    import matplotlib.pyplot as plt
except ImportError:
    ezdxf = None

# --- Finestra principale con supporto Drag and Drop ---
class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

# --- Classe Principale dell'Applicazione ---
class QuoteSimulatorApp(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.pack(fill="both", expand=True)

        # --- Stato dell'applicazione ---
        self.image_path = None
        self.original_pil_image = None
        self.unmodified_pil_image = None 
        self.display_photo_image = None
        self.canvas_image_id = None

        # --- Stato Vista (Zoom/Pan) ---
        self.view_scale = 1.0
        self.view_offset = [0, 0]
        self.is_panning = False
        self.pan_start_pos = [0, 0]
        self.zoom_job = None
        self.is_interacting = False 
        self.space_pressed = False 

        # --- Stato Modalità ---
        self.current_mode = None

        # --- Dati di Calibrazione ---
        self.calibration_mode = None
        self.scale_ratio_x = None
        self.scale_ratio_y = None
        self.global_rescale_factor = 1.0

        # Dati temporanei per la UI
        self.perspective_points = []
        self.selected_point_index = None
        self.is_drawing_line = False # --- NUOVO: Stato unificato per disegno linea
        
        # --- Stato per il ritaglio ---
        self.is_cropping = False
        self.crop_start_pos = None
        self.crop_rect_points = [] 
        self.selected_crop_corner = None

        # --- Dati di Progetto ---
        self.measurements = []

        self.fullscreen_window = None
        self.original_master = master

        # --- Lente d'ingrandimento ---
        self.loupe_window = None
        self.loupe_canvas = None
        self.loupe_photo = None
        self.loupe_size = 150
        self.loupe_zoom_var = ctk.DoubleVar(value=3.0)

        # --- Layout Principale ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.toolbar = ctk.CTkFrame(self)
        self.toolbar.grid(row=0, column=0, padx=10, pady=10, sticky="ns")

        self.canvas_container = ctk.CTkFrame(self, fg_color="#2b2b2b")
        self.canvas_container.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="nsew")
        self.canvas_container.grid_rowconfigure(0, weight=1)
        self.canvas_container.grid_columnconfigure(0, weight=1)

        self.canvas = Canvas(self.canvas_container, bg="#2b2b2b", highlightthickness=0, cursor="arrow")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        
        self.canvas.drop_target_register(DND_FILES)
        self.canvas.dnd_bind('<<Drop>>', self._handle_drop)

        self._create_toolbar_buttons()
        self._bind_events()

    def _create_toolbar_buttons(self):
        # --- MODIFICATO: Semplificata la toolbar ---
        buttons_config = [
            ("Carica File", self._load_image_dialog),
            ("Apri Progetto", self._open_project),
            ("Salva Progetto", self._save_project),
            ("---", None),
            ("Correggi Prospettiva", lambda: self._set_mode("perspective")),
            ("Ritaglia Immagine", lambda: self._set_mode("crop")),
            ("Quota", lambda: self._set_mode("quote")), # --- NUOVO: Pulsante Unico ---
            ("---", None),
            ("Ripristina Immagine", self._reset_image),
            ("Elimina Quote", self._delete_all_measurements),
            ("---", None),
            ("Esporta Immagine", self._export_image),
        ]

        for text, command in buttons_config:
            if text == "---":
                ctk.CTkFrame(self.toolbar, height=2, fg_color="gray50").pack(fill="x", padx=5, pady=5)
            else:
                btn = ctk.CTkButton(self.toolbar, text=text, command=command)
                btn.pack(padx=10, pady=5, fill="x")

        self.perspective_confirm_button = ctk.CTkButton(self.toolbar, text="✓ Conferma Prospettiva", command=self._apply_perspective_transform_to_image, fg_color="green")
        self.perspective_cancel_button = ctk.CTkButton(self.toolbar, text="✗ Annulla", command=self._cancel_perspective, fg_color="red")
        
        self.crop_confirm_button = ctk.CTkButton(self.toolbar, text="✓ Conferma Ritaglio", command=self._perform_crop, fg_color="green")
        self.crop_cancel_button = ctk.CTkButton(self.toolbar, text="✗ Annulla", command=self._cancel_crop, fg_color="red")

        ctk.CTkFrame(self.toolbar, height=2, fg_color="gray50").pack(fill="x", padx=5, pady=(15, 5))

        self.loupe_switch = ctk.CTkSwitch(self.toolbar, text="Lente", command=self._toggle_loupe)
        self.loupe_switch.pack(padx=10, pady=10, fill="x")
        self.loupe_slider = ctk.CTkSlider(self.toolbar, from_=2, to=10, variable=self.loupe_zoom_var)
        
        self.fs_button = ctk.CTkButton(self.toolbar, text="↗️ Schermo Intero", command=self._toggle_fullscreen)
        self.fs_button.pack(padx=10, pady=5, fill="x", side="bottom")
        
        self.status_label = ctk.CTkLabel(self.toolbar, text="Modalità: Nessuna", wraplength=150)
        self.status_label.pack(padx=10, pady=10, side="bottom")

    def _bind_events(self):
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<Motion>", self._on_canvas_motion)
        self.canvas.bind("<Double-Button-1>", self._on_canvas_double_click)
        
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel) 
        self.canvas.bind("<Button-4>", self._on_mouse_wheel)
        self.canvas.bind("<Button-5>", self._on_mouse_wheel)
        
        self.canvas.bind("<ButtonPress-2>", self._start_pan)
        self.canvas.bind("<B2-Motion>", self._on_pan_move)
        self.canvas.bind("<ButtonRelease-2>", self._end_pan)
        self.winfo_toplevel().bind("<KeyPress-space>", self._space_pressed)
        self.winfo_toplevel().bind("<KeyRelease-space>", self._space_released)

        self.canvas.bind("<ButtonPress-1>", self._on_point_press, add="+")
        self.canvas.bind("<B1-Motion>", self._on_point_drag, add="+")
        self.canvas.bind("<ButtonRelease-1>", self._on_point_release, add="+")

    def _set_mode(self, mode):
        self.perspective_confirm_button.pack_forget()
        self.perspective_cancel_button.pack_forget()
        self.crop_confirm_button.pack_forget()
        self.crop_cancel_button.pack_forget()

        if self.image_path is None and mode is not None:
            messagebox.showwarning("Attenzione", "Carica prima un file.", parent=self)
            return
        self.current_mode = mode
        if mode in ["perspective", "crop", "quote"]:
            self.canvas.config(cursor="cross")
            self.status_label.configure(text=f"Modalità: {mode.capitalize()}")
        else:
            self.canvas.config(cursor="arrow")
            self.status_label.configure(text="Modalità: Nessuna")
        
        if mode == "perspective": self.perspective_points = []; self.status_label.configure(text="Modalità: Prospettiva\n(Clicca 4 angoli)")
        elif mode == "crop": self.crop_rect_points = []; self.status_label.configure(text="Modalità: Ritaglia\n(Trascina per selezionare l'area)")
        # --- NUOVO: Logica di stato per la modalità Quota ---
        elif mode == "quote":
             self.is_drawing_line = False
             self.canvas.delete("temp_line")
             if self.calibration_mode:
                 self.status_label.configure(text="Modalità: Quota\n(Traccia una linea per misurare)")
             else:
                 self.status_label.configure(text="Modalità: Quota\n(Traccia una linea di riferimento per la scala)")

    def _on_canvas_click(self, event):
        if self.space_pressed: return 
        if self.current_mode == "perspective":
            if self.selected_point_index is not None: return
            self._handle_perspective_click(event)
        # --- NUOVO: Gestione unificata del click ---
        elif self.current_mode == "quote": self._handle_quote_click(event)
        elif self.current_mode == "crop": 
            if not self.is_cropping and not self.crop_rect_points:
                self._start_crop(event)

    def _on_canvas_motion(self, event):
        if self.is_drawing_line and self.canvas.find_withtag("temp_line"):
             coords = self.canvas.coords("temp_line")
             if len(coords) >= 2: self.canvas.coords("temp_line", coords[0], coords[1], event.x, event.y)
        if self.loupe_window and self.loupe_window.winfo_exists():
            self._update_loupe(event)

    def _handle_drop(self, event):
        path = event.data
        if " " in path and not (path.startswith("{") and path.endswith("}")):
            path = "{" + path + "}"
        path = self.canvas.tk.splitlist(path)[0]
        if os.path.isfile(path):
            self._process_and_load_path(path)

    def _load_image_dialog(self):
        filetypes = [("Tutti i file supportati", "*.jpg *.jpeg *.png *.bmp *.pdf *.dxf"), ("Immagini", "*.jpg *.jpeg *.png *.bmp"), ("PDF", "*.pdf"), ("DXF", "*.dxf"), ("Tutti i file", "*.*")]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path: self._process_and_load_path(path)
    
    def _convert_to_pil(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in ['.jpg', '.jpeg', '.png', '.bmp']: return Image.open(path).convert("RGBA")
        elif ext == '.pdf':
            if not fitz: messagebox.showerror("Errore Libreria", "La libreria PyMuPDF (fitz) non è installata.\nInstallala con: pip install pymupdf", parent=self); return None
            doc = fitz.open(path); page_count = len(doc); page_num = 0
            if page_count > 1:
                num = simpledialog.askinteger("Seleziona Pagina", f"Il PDF ha {page_count} pagine.\nQuale pagina vuoi importare? (1-{page_count})", initialvalue=1, minvalue=1, maxvalue=page_count, parent=self)
                if num is None: return None
                page_num = num - 1
            page = doc.load_page(page_num); pix = page.get_pixmap(dpi=200); return Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("RGBA")
        elif ext == '.dxf':
            if not ezdxf: messagebox.showerror("Errore Libreria", "La libreria ezdxf non è installata.\nInstallala con: pip install ezdxf matplotlib", parent=self); return None
            try:
                doc = ezdxf.readfile(path); msp = doc.modelspace(); fig, ax = plt.subplots(); ctx = RenderContext(doc); out = MatplotlibBackend(ax); Frontend(ctx, out).draw_layout(msp, finalize=True); buf = io.BytesIO(); fig.savefig(buf, format='png', dpi=300, bbox_inches='tight', pad_inches=0.1); plt.close(fig); buf.seek(0); return Image.open(buf).convert("RGBA")
            except Exception as e: messagebox.showerror("Errore DXF", f"Impossibile renderizzare il file DXF:\n{e}", parent=self); return None
        return None

    def _process_and_load_path(self, path):
        try:
            if not os.path.isfile(path): return
            pil_image = self._convert_to_pil(path)
            if pil_image:
                self.unmodified_pil_image = pil_image.copy(); self.original_pil_image = pil_image; self.image_path = path; self._reset_all(); self.after(50, self._fit_image_to_view) 
        except Exception as e: messagebox.showerror("Errore Caricamento", f"Impossibile caricare il file:\n{e}", parent=self)

    def _fit_image_to_view(self):
        if not self.original_pil_image: return
        canvas_w = self.canvas.winfo_width(); canvas_h = self.canvas.winfo_height(); img_w, img_h = self.original_pil_image.size
        if img_w == 0 or img_h == 0: return
        scale_w = canvas_w / img_w; scale_h = canvas_h / img_h; self.view_scale = min(scale_w, scale_h) * 0.98 
        new_w = img_w * self.view_scale; new_h = img_h * self.view_scale; self.view_offset = [(canvas_w - new_w) / 2, (canvas_h - new_h) / 2]; self._redraw_canvas()

    def _on_canvas_resize(self, event=None):
        if self.original_pil_image: self._fit_image_to_view()

    def _on_mouse_wheel(self, event):
        self.is_interacting = True
        if self.zoom_job: self.after_cancel(self.zoom_job)
        zoom_factor = 0.9 if (event.num == 5 or event.delta < 0) else 1.1; mouse_x, mouse_y = event.x, event.y; img_coord_before_zoom = self._canvas_to_image_coords((mouse_x, mouse_y)); new_scale = self.view_scale * zoom_factor
        if new_scale < 0.01 or new_scale > 50: return
        self.view_scale = new_scale; self.view_offset[0] = mouse_x - img_coord_before_zoom[0] * self.view_scale; self.view_offset[1] = mouse_y - img_coord_before_zoom[1] * self.view_scale; self._redraw_canvas(quality='low'); self.zoom_job = self.after(300, self._finalize_interaction)

    def _space_pressed(self, event): self.space_pressed = True; self.canvas.config(cursor="hand2")
    def _space_released(self, event): self.space_pressed = False; self.is_panning = False; self.canvas.config(cursor="cross" if self.current_mode in ["perspective", "crop", "quote"] else "arrow")
    def _start_pan(self, event): self.is_interacting = True; self.is_panning = True; self.pan_start_pos = [event.x, event.y]; self.canvas.config(cursor="fleur")
    def _on_pan_move(self, event):
        if not self.is_panning: return
        dx = event.x - self.pan_start_pos[0]; dy = event.y - self.pan_start_pos[1]; self.view_offset[0] += dx; self.view_offset[1] += dy; self.pan_start_pos = [event.x, event.y]; self._redraw_canvas(quality='low')
    def _end_pan(self, event): self.is_panning = False; self.canvas.config(cursor="cross" if self.current_mode else "arrow"); self._finalize_interaction()
    def _finalize_interaction(self): self.is_interacting = False; self._redraw_canvas(quality='high')

    def _handle_perspective_click(self, event):
        if len(self.perspective_points) >= 4: return
        img_point = self._canvas_to_image_coords(np.array([event.x, event.y])); self.perspective_points.append(list(img_point))
        if len(self.perspective_points) == 4:
            self.status_label.configure(text="Modalità: Prospettiva\n(Regola i punti e conferma)")
            self.perspective_confirm_button.pack(padx=10, pady=5, fill="x", after=self.toolbar.winfo_children()[5]); self.perspective_cancel_button.pack(padx=10, pady=5, fill="x", after=self.perspective_confirm_button)
        self._redraw_canvas()

    def _cancel_perspective(self): self._set_mode(None); self.perspective_points = []; self._redraw_canvas()
    def _apply_perspective_transform_to_image(self):
        try:
            def order_points(pts):
                rect = np.zeros((4, 2), dtype="float32"); s = pts.sum(axis=1); rect[0] = pts[np.argmin(s)]; rect[2] = pts[np.argmax(s)]; diff = np.diff(pts, axis=1); rect[1] = pts[np.argmin(diff)]; rect[3] = pts[np.argmax(diff)]; return rect
            src_pts = order_points(np.array(self.perspective_points, dtype="float32")); tl, tr, br, bl = src_pts; widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2)); widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2)); maxWidth = max(int(widthA), int(widthB)); heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2)); heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2)); maxHeight = max(int(heightA), int(heightB))
            if maxWidth == 0 or maxHeight == 0: return
            dst_pts = np.array([[0, 0], [maxWidth - 1, 0], [maxWidth - 1, maxHeight - 1], [0, maxHeight - 1]], dtype="float32"); matrix = cv2.getPerspectiveTransform(src_pts, dst_pts); h, w = self.original_pil_image.height, self.original_pil_image.width; original_corners = np.array([[[0, 0]], [[w - 1, 0]], [[w - 1, h - 1]], [[0, h - 1]]], dtype="float32"); transformed_corners = cv2.perspectiveTransform(original_corners, matrix); x_min, y_min = np.min(transformed_corners, axis=0).ravel().astype(int); x_max, y_max = np.max(transformed_corners, axis=0).ravel().astype(int); new_width = x_max - x_min; new_height = y_max - y_min; translation_matrix = np.array([[1, 0, -x_min], [0, 1, -y_min], [0, 0, 1]]); final_matrix = translation_matrix.dot(matrix); pil_rgba = self.original_pil_image.convert("RGBA"); cv_image = cv2.cvtColor(np.array(pil_rgba), cv2.COLOR_RGBA2BGRA); warped_image = cv2.warpPerspective(cv_image, final_matrix, (new_width, new_height), flags=cv2.INTER_LANCZOS4); corrected_pil_image = Image.fromarray(cv2.cvtColor(warped_image, cv2.COLOR_BGRA2RGBA)); self.original_pil_image = corrected_pil_image; messagebox.showinfo("Successo", "La prospettiva è stata corretta.\nOra imposta una nuova scala.", parent=self)
            self._reset_all_but_image(); self._fit_image_to_view()
        except Exception as e: messagebox.showerror("Errore Prospettiva", f"Impossibile correggere la prospettiva:\n{e}", parent=self); self._set_mode(None)

    def _on_point_press(self, event):
        if self.space_pressed: self._start_pan(event); return
        if self.current_mode == 'perspective' and len(self.perspective_points) == 4:
            canvas_coords = np.array([event.x, event.y]); min_dist = 15 
            for i, point in enumerate(self.perspective_points):
                p_canvas = self._image_to_canvas_coords(point); dist = np.linalg.norm(canvas_coords - p_canvas)
                if dist < min_dist: self.selected_point_index = i; self.canvas.config(cursor="hand2"); break
        elif self.current_mode == 'crop' and self.crop_rect_points:
            self._find_selected_crop_corner(event)

    def _on_point_drag(self, event):
        if self.space_pressed: self._on_pan_move(event); return
        if self.is_cropping:
            x0, y0 = self.crop_start_pos; x1, y1 = event.x, event.y
            self.canvas.delete("crop_rect"); self.canvas.create_rectangle(x0, y0, x1, y1, outline="red", dash=(4, 4), tags="crop_rect")
        elif self.selected_point_index is not None:
            img_coords = self._canvas_to_image_coords(np.array([event.x, event.y])); self.perspective_points[self.selected_point_index] = list(img_coords); self._redraw_canvas()
        elif self.selected_crop_corner is not None:
            self._drag_crop_corner(event)

    def _on_point_release(self, event):
        if self.space_pressed: self._end_pan(event); return
        if self.is_cropping: self._finalize_crop_selection(event)
        self.selected_point_index = None; self.selected_crop_corner = None
        if self.current_mode in ['perspective', 'crop']: self.canvas.config(cursor="cross")

    def _start_crop(self, event):
        self.is_cropping = True; self.crop_start_pos = [event.x, event.y]

    def _finalize_crop_selection(self, event):
        self.canvas.delete("crop_rect"); self.is_cropping = False
        start_canvas = self.crop_start_pos; end_canvas = [event.x, event.y]
        if abs(end_canvas[0] - start_canvas[0]) < 5 or abs(end_canvas[1] - start_canvas[1]) < 5: return
        p1_img = self._canvas_to_image_coords(start_canvas); p2_img = self._canvas_to_image_coords(end_canvas)
        self.crop_rect_points = [[min(p1_img[0], p2_img[0]), min(p1_img[1], p2_img[1])], [max(p1_img[0], p2_img[0]), max(p1_img[1], p2_img[1])]]
        self.status_label.configure(text="Modalità: Ritaglia\n(Regola gli angoli e conferma)"); self.crop_confirm_button.pack(padx=10, pady=5, fill="x", after=self.toolbar.winfo_children()[6]); self.crop_cancel_button.pack(padx=10, pady=5, fill="x", after=self.crop_confirm_button)
        self._redraw_canvas()

    def _find_selected_crop_corner(self, event):
        canvas_coords = np.array([event.x, event.y]); min_dist = 15
        p1_img, p2_img = self.crop_rect_points; corners_img = [[p1_img[0], p1_img[1]], [p2_img[0], p1_img[1]], [p2_img[0], p2_img[1]], [p1_img[0], p2_img[1]]]
        for i, corner in enumerate(corners_img):
            c_canvas = self._image_to_canvas_coords(corner); dist = np.linalg.norm(canvas_coords - c_canvas)
            if dist < min_dist: self.selected_crop_corner = i; self.canvas.config(cursor="hand2"); break
    
    def _drag_crop_corner(self, event):
        img_coords = self._canvas_to_image_coords((event.x, event.y))
        i = self.selected_crop_corner; p1, p2 = self.crop_rect_points
        if i == 0: p1[0], p1[1] = img_coords[0], img_coords[1]
        elif i == 1: p2[0], p1[1] = img_coords[0], img_coords[1]
        elif i == 2: p2[0], p2[1] = img_coords[0], img_coords[1]
        elif i == 3: p1[0], p2[1] = img_coords[0], img_coords[1]
        self.crop_rect_points = [[min(p1[0], p2[0]), min(p1[1], p2[1])], [max(p1[0], p2[0]), max(p1[1], p2[1])]]
        self._redraw_canvas()

    def _cancel_crop(self):
        self._set_mode(None); self.crop_rect_points = []; self._redraw_canvas()

    def _perform_crop(self):
        if not self.crop_rect_points: return
        p1_img, p2_img = self.crop_rect_points
        cropped_image = self.original_pil_image.crop((p1_img[0], p1_img[1], p2_img[0], p2_img[1]))
        self.original_pil_image = cropped_image
        self._reset_all_but_image(); self._fit_image_to_view()

    # --- NUOVA FUNZIONE UNIFICATA per la modalità QUOTA ---
    def _handle_quote_click(self, event):
        if not self.is_drawing_line:
            self.is_drawing_line = True
            color = "cyan" if self.calibration_mode else "yellow"
            self.canvas.create_line(event.x, event.y, event.x, event.y, fill=color, width=2, dash=(4, 2), tags="temp_line")
        else:
            p1_canvas = self.canvas.coords("temp_line")[:2]
            if not p1_canvas:
                self.is_drawing_line = False
                return

            p2_canvas = [event.x, event.y]
            self.canvas.delete("temp_line")
            self.is_drawing_line = False
            
            p1_img = self._canvas_to_image_coords(p1_canvas)
            p2_img = self._canvas_to_image_coords(p2_canvas)

            if self.calibration_mode:
                # Se siamo già calibrati, prendi una misura
                distance = self._calculate_distance(p1_img, p2_img)
                self.measurements.append({"p1_img": list(p1_img), "p2_img": list(p2_img), "distance": distance})
                self._redraw_canvas()
            else:
                # Se non siamo calibrati, questa linea imposta la scala
                self._ask_scale_dimension(p1_img, p2_img)

    def _ask_scale_dimension(self, p1_img, p2_img):
        try:
            dist_str = simpledialog.askstring("Lunghezza Reale", "Inserisci la lunghezza reale di questa linea:", parent=self)
            if dist_str is None: return
            real_dist = float(dist_str.replace(",", "."))
            dx_pix, dy_pix = np.abs(np.array(p2_img) - np.array(p1_img))
            is_horizontal = dx_pix > dy_pix
            
            is_first_ref = self.scale_ratio_x is None and self.scale_ratio_y is None

            if is_horizontal:
                self.scale_ratio_x = real_dist / dx_pix if dx_pix != 0 else float('inf')
                self.measurements = [m for m in self.measurements if not m.get("is_scale_ref_x")]
                self.measurements.append({"p1_img": list(p1_img), "p2_img": list(p2_img), "distance": real_dist, "is_scale_ref_x": True})
                if is_first_ref:
                    self.scale_ratio_y = self.scale_ratio_x
                    if messagebox.askyesno("Scala Avanzata", "Scala Larghezza impostata.\nLe proporzioni sono ora bloccate.\n\nVuoi impostare una scala separata per l'altezza (modalità avanzata)?", parent=self):
                        self.status_label.configure(text="Modalità: Quota\n(Traccia una linea di riferimento per l'altezza)")
                    else:
                        self.calibration_mode = 'simple_scale'
                        self._set_mode('quote') # Aggiorna lo stato
            else: # Verticale
                self.scale_ratio_y = real_dist / dy_pix if dy_pix != 0 else float('inf')
                self.measurements = [m for m in self.measurements if not m.get("is_scale_ref_y")]
                self.measurements.append({"p1_img": list(p1_img), "p2_img": list(p2_img), "distance": real_dist, "is_scale_ref_y": True})
                if is_first_ref:
                    self.scale_ratio_x = self.scale_ratio_y
                    if messagebox.askyesno("Scala Avanzata", "Scala Altezza impostata.\nLe proporzioni sono ora bloccate.\n\nVuoi impostare una scala separata per la larghezza (modalità avanzata)?", parent=self):
                         self.status_label.configure(text="Modalità: Quota\n(Traccia una linea di riferimento per la larghezza)")
                    else:
                        self.calibration_mode = 'simple_scale'
                        self._set_mode('quote') # Aggiorna lo stato
            
            # Se è la seconda calibrazione o l'utente ha detto no alla prima
            if not is_first_ref or not self.calibration_mode:
                self.calibration_mode = 'simple_scale'
                self._set_mode('quote')
            
            self._recalculate_all_distances()
            self._redraw_canvas()
        except (ValueError, TypeError, ZeroDivisionError): messagebox.showerror("Errore", "Valore non valido o linea di lunghezza zero.", parent=self)

    def _calculate_distance(self, p1_img, p2_img):
        if self.calibration_mode != 'simple_scale': return 0
        p1, p2 = np.array(p1_img), np.array(p2_img); dx_pix, dy_pix = np.abs(p2 - p1); ratio_x = self.scale_ratio_x if self.scale_ratio_x else self.scale_ratio_y; ratio_y = self.scale_ratio_y if self.scale_ratio_y else self.scale_ratio_x
        if not ratio_x or not ratio_y: return 0
        dx_real = dx_pix * ratio_x; dy_real = dy_pix * ratio_y; real_distance = np.sqrt(dx_real**2 + dy_real**2)
        return real_distance * self.global_rescale_factor
        
    def _on_canvas_double_click(self, event):
        mouse_pos = np.array([event.x, event.y]); target_measurement = None; min_dist_px = 15
        for m in self.measurements:
            p1_canvas = self._image_to_canvas_coords(m["p1_img"]); p2_canvas = self._image_to_canvas_coords(m["p2_img"]); line_vec = p2_canvas - p1_canvas; point_vec = mouse_pos - p1_canvas; line_len_sq = np.dot(line_vec, line_vec)
            if line_len_sq == 0: continue
            try:
                d = np.linalg.norm(np.cross(line_vec, point_vec)) / np.sqrt(line_len_sq)
                if d < min_dist_px: min_dist_px = d; target_measurement = m
            except: continue
        if target_measurement:
            new_val_str = simpledialog.askstring("Nuovo Valore", "Inserisci il nuovo valore per questa misura:", initialvalue=f"{target_measurement['distance']:.2f}", parent=self)
            if new_val_str:
                try: new_val = float(new_val_str.replace(",", ".")); self._rescale_all_measurements(target_measurement, new_val)
                except ValueError: messagebox.showerror("Errore", "Valore non valido.", parent=self)

    def _rescale_all_measurements(self, ref_measurement, new_ref_value):
        old_ref_value = ref_measurement["distance"];
        if old_ref_value == 0: return
        scale_ratio = new_ref_value / old_ref_value; self.global_rescale_factor *= scale_ratio
        for m in self.measurements: m["distance"] *= scale_ratio
        self._redraw_canvas()

    def _save_project(self):
        if not self.image_path: messagebox.showwarning("Attenzione", "Nessuna immagine caricata.", parent=self); return
        save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Progetto Simulazione Quote", "*.json")], initialdir=os.path.dirname(self.image_path))
        if not save_path: return
        project_data = {"image_path": os.path.abspath(self.image_path), "calibration_mode": self.calibration_mode, "scale_ratio_x": self.scale_ratio_x, "scale_ratio_y": self.scale_ratio_y, "global_rescale_factor": self.global_rescale_factor, "measurements": self.measurements}
        try:
            with open(save_path, "w") as f: json.dump(project_data, f, indent=4); messagebox.showinfo("Successo", "Progetto salvato con successo.", parent=self)
        except Exception as e: messagebox.showerror("Errore", f"Impossibile salvare il progetto:\n{e}", parent=self)

    def _open_project(self):
        open_path = filedialog.askopenfilename(filetypes=[("Progetto Simulazione Quote", "*.json")])
        if not open_path: return
        try:
            with open(open_path, "r") as f: project_data = json.load(f)
            image_path_from_project = project_data["image_path"]
            if not os.path.exists(image_path_from_project): raise FileNotFoundError(f"Il file immagine originale non è stato trovato:\n{image_path_from_project}")
            pil_image = self._convert_to_pil(image_path_from_project)
            if pil_image:
                self.unmodified_pil_image = pil_image.copy(); self.image_path = image_path_from_project; self._process_and_load_path(image_path_from_project) 
                self.calibration_mode = project_data.get("calibration_mode"); self.scale_ratio_x = project_data.get("scale_ratio_x"); self.scale_ratio_y = project_data.get("scale_ratio_y"); self.global_rescale_factor = project_data.get("global_rescale_factor", 1.0); self.measurements = project_data.get("measurements", [])
                self.after(100, self._redraw_canvas)
        except Exception as e: messagebox.showerror("Errore", f"Impossibile aprire il progetto:\n{e}", parent=self)

    def _export_image(self):
        if not self.original_pil_image: return
        filetypes = [("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("PDF Document", "*.pdf")]
        export_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=filetypes, initialdir=os.path.dirname(self.image_path) if self.image_path else ".")
        if not export_path: return
        
        img_to_export = self.original_pil_image.copy().convert("RGBA"); draw = ImageDraw.Draw(img_to_export)
        try: font = ImageFont.truetype("arial.ttf", 40)
        except IOError: font = ImageFont.load_default()
        for m in self.measurements:
            p1_img, p2_img = m["p1_img"], m["p2_img"]; color = "cyan"
            if m.get("is_scale_ref_x"): color = "yellow"
            if m.get("is_scale_ref_y"): color = "orange"
            draw.line([tuple(p1_img), tuple(p2_img)], fill=color, width=5); mid_x = (p1_img[0] + p2_img[0]) / 2; mid_y = (p1_img[1] + p2_img[1]) / 2
            text = f"{m['distance']:.2f}"
            text_bbox = draw.textbbox((mid_x, mid_y), text, font=font, anchor="ms")
            draw.rounded_rectangle(text_bbox, fill="#404040", radius=5)
            draw.text((mid_x, mid_y), text, fill="white", font=font, anchor="ms")
        
        if export_path.lower().endswith('.pdf'):
            try:
                A4_SIZE_PX = (2480, 3508); pdf_page = Image.new('RGB', A4_SIZE_PX, 'white'); img_to_export_rgb = img_to_export.convert('RGB')
                img_to_export_rgb.thumbnail((A4_SIZE_PX[0] - 200, A4_SIZE_PX[1] - 200), Image.Resampling.LANCZOS)
                paste_x = (A4_SIZE_PX[0] - img_to_export_rgb.width) // 2; paste_y = (A4_SIZE_PX[1] - img_to_export_rgb.height) // 2
                pdf_page.paste(img_to_export_rgb, (paste_x, paste_y))
                pdf_page.save(export_path, "PDF", resolution=300.0)
            except Exception as e: messagebox.showerror("Errore PDF", f"Impossibile salvare il PDF:\n{e}", parent=self); return
        else:
            img_to_export.save(export_path)

        messagebox.showinfo("Successo", f"File esportato in:\n{export_path}", parent=self)

    def _toggle_fullscreen(self):
        if self.fullscreen_window is None:
            self.fullscreen_window = ctk.CTkToplevel(self); self.fullscreen_window.attributes("-fullscreen", True); self.fullscreen_window.bind("<Escape>", lambda e: self._toggle_fullscreen()); self.toolbar.pack(in_=self.fullscreen_window, side="left", fill="y", padx=10, pady=10); self.canvas_container.pack(in_=self.fullscreen_window, side="left", fill="both", expand=True, padx=(0,10), pady=10); self.fs_button.configure(text="↙️ Esci Schermo Intero")
        else:
            self.toolbar.pack(in_=self, side="left", fill="y", padx=10, pady=10); self.canvas_container.pack(in_=self, side="left", fill="both", expand=True, padx=(0,10), pady=10); self.fs_button.configure(text="↗️ Schermo Intero"); self.fullscreen_window.destroy(); self.fullscreen_window = None

    def _canvas_to_image_coords(self, canvas_coords): return (np.array(canvas_coords) - self.view_offset) / self.view_scale
    def _image_to_canvas_coords(self, image_coords): return (np.array(image_coords) * self.view_scale) + self.view_offset
    
    def _reset_all_but_image(self): self._reset_calibration(); self.measurements = []; self.crop_rect_points = []; self._redraw_canvas()
    def _reset_image(self): 
        if self.unmodified_pil_image: self.original_pil_image = self.unmodified_pil_image.copy(); self._reset_all_but_image(); self._fit_image_to_view()
    def _delete_all_measurements(self): self.measurements = []; self._redraw_canvas()
    def _recalculate_all_distances(self): 
        for m in self.measurements:
            if not m.get("is_scale_ref_x") and not m.get("is_scale_ref_y"): m["distance"] = self._calculate_distance(m["p1_img"], m["p2_img"])

    def _reset_all(self): self.canvas.delete("all"); self.canvas_image_id = None; self._reset_calibration(); self.measurements = []; self._reset_view(); self._set_mode(None)
    def _reset_view(self): self.view_scale = 1.0; self.view_offset = [0, 0]
    def _reset_calibration(self): self.calibration_mode = None; self.scale_ratio_x = None; self.scale_ratio_y = None; self.global_rescale_factor = 1.0; self.perspective_points = []; self.is_drawing_line = False; self._set_mode(None)
    
    def _toggle_loupe(self): 
        if self.loupe_switch.get() == 1:
            if not self.loupe_window or not self.loupe_window.winfo_exists():
                self.loupe_window = ctk.CTkToplevel(self); self.loupe_window.overrideredirect(True); self.loupe_window.attributes("-topmost", True); self.loupe_window.geometry(f"{self.loupe_size}x{self.loupe_size}")
                self.loupe_window.attributes("-transparentcolor", "black")
                self.loupe_canvas = Canvas(self.loupe_window, width=self.loupe_size, height=self.loupe_size, highlightthickness=0, bg="black"); self.loupe_canvas.pack()
            self.loupe_slider.pack(padx=10, pady=5, fill="x")
        else:
            if self.loupe_window: self.loupe_window.destroy(); self.loupe_window = None
            self.loupe_slider.pack_forget()

    def _update_loupe(self, event): 
        if not self.original_pil_image: return
        zoom = self.loupe_zoom_var.get(); img_x, img_y = self._canvas_to_image_coords((event.x, event.y)); half_size = self.loupe_size / (2 * zoom); box = (img_x - half_size, img_y - half_size, img_x + half_size, img_y + half_size)
        try:
            region = self.original_pil_image.crop(box); magnified_region = region.resize((self.loupe_size, self.loupe_size), Image.Resampling.NEAREST)
            mask = Image.new('L', (self.loupe_size, self.loupe_size), 0); draw = ImageDraw.Draw(mask); draw.ellipse((0, 0, self.loupe_size, self.loupe_size), fill=255)
            magnified_region.putalpha(mask)
            self.loupe_photo = ImageTk.PhotoImage(magnified_region); self.loupe_canvas.create_image(0, 0, anchor="nw", image=self.loupe_photo)
            self.loupe_canvas.create_line(self.loupe_size/2, 0, self.loupe_size/2, self.loupe_size, fill="red"); self.loupe_canvas.create_line(0, self.loupe_size/2, self.loupe_size, self.loupe_size/2, fill="red"); self.loupe_window.geometry(f"+{event.x_root + 20}+{event.y_root + 20}")
        except Exception: pass

    def _redraw_canvas(self, quality='high'):
        if not self.original_pil_image: self.canvas.delete("all"); return
        resampling_method = Image.Resampling.NEAREST if quality == 'low' else Image.Resampling.LANCZOS
        self.canvas.delete("all"); img_w, img_h = self.original_pil_image.size; scaled_w, scaled_h = int(img_w * self.view_scale), int(img_h * self.view_scale)
        if scaled_w > 0 and scaled_h > 0:
            try:
                resized_img = self.original_pil_image.resize((scaled_w, scaled_h), resampling_method); self.display_photo_image = ImageTk.PhotoImage(resized_img); self.canvas.create_image(self.view_offset[0], self.view_offset[1], anchor="nw", image=self.display_photo_image)
            except Exception as e: print(f"Error during image redraw: {e}")
        if self.is_interacting: return
        if self.current_mode == "perspective" and len(self.perspective_points) > 0:
            canvas_points = [tuple(self._image_to_canvas_coords(p)) for p in self.perspective_points]
            if len(canvas_points) > 1: self.canvas.create_polygon(canvas_points, fill="red", outline="red", stipple="gray25", width=2, dash=(4,4))
            for i, p in enumerate(canvas_points): fill_color = "yellow" if i == self.selected_point_index else "red"; self.canvas.create_oval(p[0]-5, p[1]-5, p[0]+5, p[1]+5, fill=fill_color, outline="white", width=2)
        
        if self.current_mode == "crop" and self.crop_rect_points:
            p1_img, p2_img = self.crop_rect_points; p1_canvas = self._image_to_canvas_coords(p1_img); p2_canvas = self._image_to_canvas_coords(p2_img)
            self.canvas.create_rectangle(tuple(p1_canvas), tuple(p2_canvas), outline="red", dash=(4,4), width=2)
            corners_img = [[p1_img[0], p1_img[1]], [p2_img[0], p1_img[1]], [p2_img[0], p2_img[1]], [p1_img[0], p2_img[1]]]
            for i, p_img in enumerate(corners_img):
                p_canvas = self._image_to_canvas_coords(p_img); fill_color = "yellow" if i == self.selected_crop_corner else "red"
                self.canvas.create_oval(p_canvas[0]-5, p_canvas[1]-5, p_canvas[0]+5, p_canvas[1]+5, fill=fill_color, outline="white", width=2)

        for m in self.measurements:
            p1_canvas = self._image_to_canvas_coords(m["p1_img"]); p2_canvas = self._image_to_canvas_coords(m["p2_img"]); color = "cyan"
            if m.get("is_scale_ref_x"): color = "yellow"
            if m.get("is_scale_ref_y"): color = "orange"
            self.canvas.create_line(tuple(p1_canvas), tuple(p2_canvas), fill=color, width=3); mid_point = (p1_canvas + p2_canvas) / 2
            
            text_id = self.canvas.create_text(mid_point[0], mid_point[1] - 10, text=f"{m['distance']:.2f}", fill="white", font=("Arial", 10, "bold"))
            bbox = self.canvas.bbox(text_id)
            if bbox:
                padding = 4
                rect_bbox = (bbox[0] - padding, bbox[1] - padding, bbox[2] + padding, bbox[3] + padding)
                rect_id = self.canvas.create_rectangle(rect_bbox, fill="#404040", outline="")
                self.canvas.lower(rect_id, text_id)
            
def create_tab(tab_view):
    tab_name = "Simulazione quote"
    try: tab_view.dnd_bind 
    except AttributeError: print("Warning: Container does not seem to support TkinterDnD.")
    tab = tab_view.add(tab_name); app_instance = QuoteSimulatorApp(master=tab); return tab_name, app_instance

# --- Per testing standalone ---
if __name__ == '__main__':
    root = App(); root.title("Simulazione Quote Standalone"); root.geometry("1280x720"); app = QuoteSimulatorApp(master=root); root.mainloop()

