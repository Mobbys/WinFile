# app_liste_anteprime.py - v4.7 (PySide6) - Correzione freccia tema scuro
import os
import traceback
import webbrowser
import tempfile
import base64
import csv
import html
import io

# Import specifici per la gestione avanzata degli appunti su Windows
if os.name == 'nt':
    import ctypes
    from ctypes import wintypes

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                               QTreeWidget, QTreeWidgetItem, QLabel, QFileDialog,
                               QMessageBox, QHeaderView, QSplitter, QMenu, QDialog,
                               QRadioButton, QSpinBox, QDialogButtonBox, QTreeWidgetItemIterator,
                               QStyleFactory) # Aggiunto QStyleFactory
from PySide6.QtCore import Qt, Signal, Slot, QObject, QThread, QMimeData
from PySide6.QtGui import QPixmap, QImage, QGuiApplication
from PIL import Image
import fitz  # PyMuPDF

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4, landscape
from collections import defaultdict

Image.MAX_IMAGE_PIXELS = None
SUPPORTED_EXTENSIONS = ('.jpg', '.jpeg', '.tif', '.tiff', '.pdf', '.ai')
DEFAULT_DPI = 96

class ExportOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Opzioni di Esportazione PDF")
        self.layout = QVBoxLayout(self)
        orientation_layout = QHBoxLayout()
        orientation_layout.addWidget(QLabel("Orientamento Pagina:"))
        self.portrait_rb = QRadioButton("Verticale")
        self.landscape_rb = QRadioButton("Orizzontale")
        self.portrait_rb.setChecked(True)
        orientation_layout.addWidget(self.portrait_rb)
        orientation_layout.addWidget(self.landscape_rb)
        self.layout.addLayout(orientation_layout)
        columns_layout = QHBoxLayout()
        columns_layout.addWidget(QLabel("Colonne per Riga:"))
        self.columns_spinbox = QSpinBox()
        self.columns_spinbox.setRange(1, 20)
        self.columns_spinbox.setValue(4)
        columns_layout.addWidget(self.columns_spinbox)
        self.layout.addLayout(columns_layout)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)

    def get_options(self):
        return {
            "orientation": "portrait" if self.portrait_rb.isChecked() else "landscape",
            "columns": self.columns_spinbox.value()
        }

class Worker(QObject):
    finished = Signal(object, bool)
    
    def __init__(self, task_function, *args, **kwargs):
        super().__init__()
        self.task_function = task_function
        self.args = args
        self.kwargs = kwargs

    @Slot()
    def run(self):
        try:
            result = self.task_function(*self.args, **self.kwargs)
            self.finished.emit(result, True)
        except Exception as e:
            traceback.print_exc()
            self.finished.emit(str(e), False)

class FileScannerApp(QWidget):
    def __init__(self, master, winfile_app):
        super().__init__(master)
        self.winfile_app = winfile_app
        self.scan_results = {}
        self.is_scanning = False
        
        self.create_widgets()
        self.create_layouts()
        self.create_connections()
        
    def create_widgets(self):
        self.select_button = QPushButton("Aggiungi Cartella")
        self.clear_button = QPushButton("Svuota Lista")
        
        self.copy_all_button = QPushButton("Copia Tabella")
        self.copy_formatted_button = QPushButton("Copia Formattata")
        self.print_button = QPushButton("Stampa Tabella")
        
        self.export_html_button = QPushButton("Anteprima HTML")
        self.export_pdf_button = QPushButton("Esporta in PDF")
        
        self.clear_button.setEnabled(False)
        self.copy_all_button.setEnabled(False)
        self.copy_formatted_button.setEnabled(False)
        self.print_button.setEnabled(False)
        self.export_html_button.setEnabled(False)
        self.export_pdf_button.setEnabled(False)
        
        self.tree = QTreeWidget()
        # --- MODIFICA CHIAVE ---
        # Applica lo stile "Fusion" solo a questo widget per garantire la visibilità
        # delle frecce di espansione nel tema scuro.
        self.tree.setStyle(QStyleFactory.create('Fusion'))
        
        self.tree.setColumnCount(3)
        self.tree.setHeaderLabels(["Nome File / Pagina", "Dimensioni (cm)", "Percorso"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.header().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)

        self.preview_label = QLabel("Trascina i file o le cartelle qui")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumSize(300, 300)

        self.status_label = QLabel("Pronto.")

    def create_layouts(self):
        self.main_layout = QVBoxLayout(self)
        self.top_layout = QHBoxLayout()
        self.top_layout.addWidget(self.select_button)
        self.top_layout.addWidget(self.clear_button)
        self.top_layout.addWidget(self.copy_all_button)
        self.top_layout.addWidget(self.copy_formatted_button)
        self.top_layout.addWidget(self.print_button)
        self.top_layout.addWidget(self.export_html_button)
        self.top_layout.addWidget(self.export_pdf_button)
        self.top_layout.addStretch()
        
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.addWidget(self.tree)
        self.splitter.addWidget(self.preview_label)
        self.splitter.setSizes([700, 300])
        self.main_layout.addLayout(self.top_layout)
        self.main_layout.addWidget(self.splitter, 1)
        self.main_layout.addWidget(self.status_label)

    def create_connections(self):
        self.select_button.clicked.connect(self.select_folder_dialog)
        self.clear_button.clicked.connect(self.clear_results)
        self.copy_all_button.clicked.connect(lambda: self.copy_to_clipboard(as_html=False, for_context_menu=False))
        self.copy_formatted_button.clicked.connect(lambda: self.copy_to_clipboard(as_html=True, for_context_menu=False))
        self.print_button.clicked.connect(lambda: self.print_selection(for_context_menu=False))
        self.export_html_button.clicked.connect(lambda: self.export_to_html(for_context_menu=False))
        self.export_pdf_button.clicked.connect(lambda: self.export_to_pdf(for_context_menu=False))
        self.tree.currentItemChanged.connect(self.on_item_select)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)

    def handle_drop_event(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            self.run_scan(paths)

    def show_context_menu(self, position):
        if not self.tree.selectedItems(): return
        menu = QMenu()
        menu.addAction("Copia selezione", lambda: self.copy_to_clipboard(as_html=False, for_context_menu=True))
        menu.addAction("Copia selezione formattata", lambda: self.copy_to_clipboard(as_html=True, for_context_menu=True))
        menu.addAction("Stampa selezione...", lambda: self.print_selection(for_context_menu=True))
        menu.addSeparator()
        menu.addAction("Esporta selezione in CSV...", lambda: self.export_to_csv(for_context_menu=True))
        menu.addAction("Crea Anteprima HTML selezione...", lambda: self.export_to_html(for_context_menu=True))
        menu.addAction("Esporta selezione in PDF...", lambda: self.export_to_pdf(for_context_menu=True))
        menu.addSeparator()
        menu.addAction("Elimina selezione dalla lista", self.remove_selected_items)
        menu.exec(self.tree.viewport().mapToGlobal(position))

    def remove_selected_items(self):
        """Rimuove i file selezionati (o i file a cui appartengono le pagine selezionate) dalla lista."""
        selected_items = self.tree.selectedItems()
        if not selected_items:
            return

        # Identifica i file unici associati alla selezione
        files_to_remove = set()
        for item in selected_items:
            parent = item.parent()
            if parent:  # È una pagina, quindi risale al file genitore
                full_path = parent.data(0, Qt.ItemDataRole.UserRole)
            else:  # È già un file
                full_path = item.data(0, Qt.ItemDataRole.UserRole)
            
            if full_path:
                files_to_remove.add(full_path)

        if not files_to_remove:
            return

        reply = QMessageBox.question(self, 'Conferma Eliminazione',
                                     f"Questo rimuoverà {len(files_to_remove)} file (e tutte le loro pagine) dalla lista. Continuare?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.No:
            return

        # Trova e rimuove gli item di primo livello corrispondenti ai file da eliminare
        items_to_take = []
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if item and item.data(0, Qt.ItemDataRole.UserRole) in files_to_remove:
                items_to_take.append(item)

        for item in items_to_take:
            index = self.tree.indexOfTopLevelItem(item)
            self.tree.takeTopLevelItem(index)

        # Pulisce il dizionario dei dati
        for path in files_to_remove:
            if path in self.scan_results:
                del self.scan_results[path]
                
        self.status_label.setText(f"{len(files_to_remove)} file rimossi dalla lista.")
        
        if not self.scan_results:
            # Disabilita i pulsanti se la lista è vuota
            for button in [self.clear_button, self.copy_all_button, self.copy_formatted_button, self.print_button, self.export_html_button, self.export_pdf_button]:
                button.setEnabled(False)

    def run_scan(self, paths):
        if self.is_scanning: return
        self.is_scanning = True
        self.status_label.setText("Scansione in corso...")
        self.select_button.setEnabled(False)
        self.run_task_in_thread(self._scan_task, self.on_scan_finished, paths)

    def run_task_in_thread(self, task_function, on_finish_slot, *args, **kwargs):
        self.thread = QThread()
        self.worker = Worker(task_function, *args, **kwargs)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(on_finish_slot)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def _scan_task(self, paths):
        results = []
        for path in paths:
            scan_root = os.path.normpath(path)
            if os.path.isdir(scan_root):
                for root, _, files in os.walk(scan_root):
                    for f in files:
                        if f.lower().endswith(SUPPORTED_EXTENSIONS):
                            details = self._get_file_details(os.path.join(root, f))
                            if details: details['scan_root'] = scan_root; results.append(details)
            elif os.path.isfile(scan_root) and scan_root.lower().endswith(SUPPORTED_EXTENSIONS):
                details = self._get_file_details(scan_root)
                if details: details['scan_root'] = os.path.dirname(scan_root); results.append(details)
        return results

    @Slot(object, bool)
    def on_scan_finished(self, new_files, success):
        if success: self.add_scan_results(new_files)
        else:
            self.status_label.setText("Errore durante la scansione.")
            QMessageBox.critical(self, "Errore di Scansione", f"Si è verificato un errore:\n{new_files}")
        self.is_scanning = False
        self.select_button.setEnabled(True)

    def add_scan_results(self, new_files):
        for item_data in new_files:
            full_path = item_data['full_path']
            if full_path not in self.scan_results:
                self.scan_results[full_path] = item_data
                parent_item = QTreeWidgetItem(self.tree)
                parent_item.setText(0, item_data['filename'])
                parent_item.setText(2, self._get_display_path(item_data))
                parent_item.setData(0, Qt.ItemDataRole.UserRole, full_path)
                if item_data['page_count'] > 1:
                    for i, page_detail in enumerate(item_data['pages_details']):
                        child_item = QTreeWidgetItem(parent_item); child_item.setText(0, f"  Pagina {i+1}"); child_item.setText(1, page_detail['dimensions_cm'])
                else: parent_item.setText(1, item_data['pages_details'][0]['dimensions_cm'])
        has_results = bool(self.scan_results)
        self.status_label.setText(f"Scansione completata. Trovati {len(self.scan_results)} file.")
        for button in [self.clear_button, self.copy_all_button, self.copy_formatted_button, self.print_button, self.export_html_button, self.export_pdf_button]:
            button.setEnabled(has_results)

    def select_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleziona Cartella")
        if folder: self.run_scan([folder])

    def clear_results(self):
        self.tree.clear(); self.scan_results.clear()
        self.preview_label.clear(); self.preview_label.setText("Trascina i file o le cartelle qui")
        self.status_label.setText("Pronto.")
        for button in [self.clear_button, self.copy_all_button, self.copy_formatted_button, self.print_button, self.export_html_button, self.export_pdf_button]:
            button.setEnabled(False)

    @Slot(QTreeWidgetItem, QTreeWidgetItem)
    def on_item_select(self, current, previous):
        if not current: return
        item_to_load = current if current.parent() is None else current.parent()
        page_index = 0 if current.parent() is None else item_to_load.indexOfChild(current)
        full_path = item_to_load.data(0, Qt.ItemDataRole.UserRole)
        try:
            ext = os.path.splitext(full_path)[1].lower()
            pixmap = None
            if ext in ('.pdf', '.ai'):
                doc = fitz.open(full_path); page = doc.load_page(page_index); pix = page.get_pixmap()
                image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).rgbSwapped()
                pixmap = QPixmap.fromImage(image)
            else: pixmap = QPixmap(full_path)
            if pixmap and not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(self.preview_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                self.preview_label.setPixmap(scaled_pixmap)
            else: self.preview_label.setText("Anteprima non disponibile")
        except Exception as e:
            self.preview_label.setText("Anteprima non disponibile"); print(f"Errore anteprima: {e}")

    def _get_items_to_process(self, for_context_menu):
        selected_items = self.tree.selectedItems()
        if for_context_menu:
            return selected_items
        else:
            if selected_items:
                return selected_items
            else:
                all_items = []
                it = QTreeWidgetItemIterator(self.tree)
                while it.value():
                    all_items.append(it.value())
                    it += 1
                return all_items

    def copy_to_clipboard(self, as_html=False, for_context_menu=False):
        items = self._get_items_to_process(for_context_menu)
        if not items:
            self.status_label.setText("Nessun elemento selezionato da copiare.")
            return

        header = ["Nome File / Pagina", "Dimensioni (cm)", "Percorso", "Area (m²)"]
        rows = []
        total_area = 0

        for item in items:
            is_child = item.parent() is not None
            parent_item = item.parent() if is_child else item
            full_path = parent_item.data(0, Qt.ItemDataRole.UserRole)
            if not full_path or full_path not in self.scan_results: continue
            file_info = self.scan_results[full_path]
            page_index = parent_item.indexOfChild(item) if is_child else 0
            
            if 0 <= page_index < len(file_info['pages_details']):
                page_details = file_info['pages_details'][page_index]
                width_cm = page_details.get('width_cm', 0)
                height_cm = page_details.get('height_cm', 0)
                area_sqm = (width_cm * height_cm) / 10000
                total_area += area_sqm
                rows.append([
                    item.text(0).strip(),
                    page_details.get('dimensions_cm', 'N/D'),
                    self._get_display_path(file_info),
                    f"{area_sqm:.4f}"
                ])

        if not rows:
            self.status_label.setText("Nessun dato valido trovato per la copia.")
            return

        if not as_html:
            plain_text_lines = ["\t".join(header)] + ["\t".join(map(str, row)) for row in rows]
            QGuiApplication.clipboard().setText("\n".join(plain_text_lines))
        else:
            # Stili in linea per massima compatibilità
            table_html = '<table border="1" style="border-collapse: collapse; width: 100%; font-family: sans-serif;">'
            table_html += '<thead style="background-color: #e0e0e0;"><tr>'
            for h in header:
                table_html += f'<th style="padding: 5px; text-align: left;">{html.escape(h)}</th>'
            table_html += '</tr></thead><tbody>'
            for row in rows:
                table_html += '<tr>'
                for i, cell in enumerate(row):
                    style = 'padding: 5px; text-align: left;'
                    if i == 0 and str(cell).startswith("Pagina"):
                         style = 'padding: 5px; padding-left: 20px; text-align: left;'
                    table_html += f'<td style="{style}">{html.escape(str(cell))}</td>'
                table_html += '</tr>'
            table_html += '</tbody>'
            table_html += '<tfoot style="font-weight: bold; background-color: #f0f0f0;"><tr>'
            table_html += f'<td style="padding: 5px;" colspan="3">Totale ({len(rows)} elementi)</td>'
            table_html += f'<td style="padding: 5px;">{total_area:.4f} m²</td>'
            table_html += '</tr></tfoot></table>'

            try:
                if os.name == 'nt':
                    self._set_clipboard_html_windows(table_html)
                else:
                    # Fallback per sistemi non-Windows
                    mime_data = QMimeData()
                    mime_data.setHtml(table_html)
                    QGuiApplication.clipboard().setMimeData(mime_data)
                self.status_label.setText(f"Tabella formattata ({len(rows)} righe) copiata.")
            except Exception as e:
                self.status_label.setText("Errore durante la copia formattata.")
                QMessageBox.critical(self, "Errore Appunti", f"Impossibile copiare l'HTML.\nDettagli: {e}")
                traceback.print_exc()

    def _set_clipboard_html_windows(self, html_fragment: str):
        """Mette una stringa HTML negli appunti di Windows usando ctypes."""
        if os.name != 'nt':
            raise NotImplementedError("Questa funzione è solo per Windows.")

        user32 = ctypes.WinDLL('user32')
        kernel32 = ctypes.WinDLL('kernel32')
        wintypes.HGLOBAL = wintypes.HANDLE
        user32.OpenClipboard.argtypes = [wintypes.HWND]
        user32.OpenClipboard.restype = wintypes.BOOL
        user32.CloseClipboard.restype = wintypes.BOOL
        user32.EmptyClipboard.restype = wintypes.BOOL
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HGLOBAL]
        user32.SetClipboardData.restype = wintypes.HGLOBAL
        user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
        user32.RegisterClipboardFormatW.restype = wintypes.UINT
        kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = wintypes.LPVOID
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalUnlock.restype = wintypes.BOOL
        kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]

        GMEM_MOVEABLE = 0x0002
        CF_HTML = user32.RegisterClipboardFormatW("HTML Format")

        header_template = (
            "Version:0.9\r\n"
            "StartHTML:{{:0>9}}\r\n"
            "EndHTML:{{:0>9}}\r\n"
            "StartFragment:{{:0>9}}\r\n"
            "EndFragment:{{:0>9}}\r\n"
        )
        html_body_template = (
            "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>"
            "<!--StartFragment-->{}<!--EndFragment-->"
            "</body></html>"
        )
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
            try:
                ctypes.memmove(p_global_mem, clipboard_data, len(clipboard_data))
            finally:
                kernel32.GlobalUnlock(h_global_mem)
            if not user32.SetClipboardData(CF_HTML, h_global_mem):
                raise ctypes.WinError()
            h_global_mem = None
        finally:
            user32.CloseClipboard()
            if h_global_mem:
                kernel32.GlobalFree(h_global_mem)

    def export_to_csv(self, for_context_menu=False):
        items = self._get_items_to_process(for_context_menu)
        if not items: return
        
        path, _ = QFileDialog.getSaveFileName(self, "Salva come CSV", "", "CSV Files (*.csv)")
        if not path: return

        header = ["Nome File / Pagina", "Dimensioni (cm)", "Percorso", "Area (m²)"]
        rows = []
        for item in items:
            path = item.data(0, Qt.ItemDataRole.UserRole)
            is_child = item.parent() is not None
            if is_child:
                path = item.parent().data(0, Qt.ItemDataRole.UserRole)

            file_info = self.scan_results.get(path)
            if not file_info: continue
            
            page_num = item.parent().indexOfChild(item) if is_child else 0
            if page_num < len(file_info['pages_details']):
                page_details = file_info['pages_details'][page_num]
                area_sqm = (page_details.get('width_cm', 0) * page_details.get('height_cm', 0)) / 10000
                row_data = [item.text(0).strip(), page_details['dimensions_cm'], file_info['path'], f"{area_sqm:.4f}"]
                rows.append(row_data)

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(header)
                writer.writerows(rows)
            self.status_label.setText(f"CSV esportato con successo in:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Errore Esportazione", f"Impossibile salvare il file CSV:\n{e}")

    def print_selection(self, for_context_menu=False):
        items = self._get_items_to_process(for_context_menu)
        if not items: return
        
        self.status_label.setText("Creazione Anteprima di Stampa in corso...")
        
        header, rows = self._get_table_data_from_items(items)
        html_content = self._generate_table_html_content(header, rows)
        
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.html', encoding='utf-8') as f:
                f.write(html_content)
            webbrowser.open(f'file://{os.path.realpath(f.name)}')
            self.status_label.setText("Anteprima di Stampa creata con successo.")
        except Exception as e:
            self.status_label.setText("Errore nella creazione dell'anteprima.")
            QMessageBox.critical(self, "Errore Stampa", f"Impossibile generare il file di stampa:\n{e}")

    def _get_table_data_from_items(self, items):
        header = ["Nome File / Pagina", "Dimensioni (cm)", "Percorso", "Area (m²)"]
        rows = []
        for item in items:
            path = item.data(0, Qt.ItemDataRole.UserRole)
            is_child = item.parent() is not None
            if is_child:
                path = item.parent().data(0, Qt.ItemDataRole.UserRole)

            file_info = self.scan_results.get(path)
            if not file_info: continue
            
            page_num = item.parent().indexOfChild(item) if is_child else 0
            if page_num < len(file_info['pages_details']):
                page_details = file_info['pages_details'][page_num]
                area_sqm = (page_details.get('width_cm', 0) * page_details.get('height_cm', 0)) / 10000
                row_data = [item.text(0).strip(), page_details['dimensions_cm'], file_info['path'], f"{area_sqm:.4f}"]
                rows.append(row_data)
        return header, rows

    def _generate_table_html_content(self, header, rows):
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='UTF-8'>
            <title>Stampa Tabella</title>
            <style>
                body {{ font-family: sans-serif; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; word-break: break-all; }}
                th {{ background-color: #f2f2f2; }}
                @media print {{
                    body {{ font-size: 10pt; }}
                    .no-print {{ display: none; }}
                }}
            </style>
        </head>
        <body onload="window.print()">
            <div class="no-print">
                <h1>Anteprima di Stampa</h1>
                <p>Se la stampa non si avvia automaticamente, premi Ctrl+P.</p>
            </div>
            <table>
                <thead>
                    <tr>{''.join(f'<th>{h}</th>' for h in header)}</tr>
                </thead>
                <tbody>
        """
        for row in rows:
            html_content += '<tr>' + ''.join(f'<td>{html.escape(str(d))}</td>' for d in row) + '</tr>'

        html_content += "</tbody></table></body></html>"
        return html_content

    def export_to_html(self, for_context_menu=False, is_print=False):
        items = self._get_items_to_process(for_context_menu)
        if not items: return

        selected_pages_map = defaultdict(set)
        for item in items:
            is_child = item.parent() is not None
            parent_item = item.parent() if is_child else item
            full_path = parent_item.data(0, Qt.ItemDataRole.UserRole)
            if not full_path or full_path not in self.scan_results: continue
            file_info = self.scan_results[full_path]
            if is_child:
                page_index = parent_item.indexOfChild(item)
                selected_pages_map[full_path].add(page_index)
            else:
                for i in range(file_info.get('page_count', 1)):
                    selected_pages_map[full_path].add(i)

        data_to_export = []
        for full_path, page_indices in selected_pages_map.items():
            data_to_export.append({
                "file_info": self.scan_results[full_path],
                "selected_pages": sorted(list(page_indices))
            })

        if not data_to_export: return
        
        self.status_label.setText("Creazione Anteprima HTML in corso...")
        self.run_task_in_thread(self._html_task, self.on_html_preview_finished, data_to_export, is_print)

    def _html_task(self, data, is_print):
        html_content = self._generate_html_content(data, is_print)
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.html', encoding='utf-8') as f:
            f.write(html_content)
        webbrowser.open(f'file://{os.path.realpath(f.name)}')
        return "Anteprima HTML creata con successo."

    def export_to_pdf(self, for_context_menu=False):
        items = self._get_items_to_process(for_context_menu)
        if not items: return

        selected_pages_map = defaultdict(set)
        for item in items:
            is_child = item.parent() is not None
            parent_item = item.parent() if is_child else item
            full_path = parent_item.data(0, Qt.ItemDataRole.UserRole)
            if not full_path or full_path not in self.scan_results: continue
            file_info = self.scan_results[full_path]
            if is_child:
                page_index = parent_item.indexOfChild(item)
                selected_pages_map[full_path].add(page_index)
            else:
                for i in range(file_info.get('page_count', 1)):
                    selected_pages_map[full_path].add(i)

        data_to_export = []
        for full_path, page_indices in selected_pages_map.items():
            data_to_export.append({
                "file_info": self.scan_results[full_path],
                "selected_pages": sorted(list(page_indices))
            })

        if not data_to_export: return

        dialog = ExportOptionsDialog(self)
        if dialog.exec():
            options = dialog.get_options()
            path, _ = QFileDialog.getSaveFileName(self, "Salva come PDF", "", "PDF Files (*.pdf)")
            if path:
                self.status_label.setText("Creazione del PDF in corso...")
                self.run_task_in_thread(self._pdf_task, self.on_export_finished, path, options, data_to_export)

    def _pdf_task(self, path, options, data):
        grouped_results = defaultdict(list)
        for item in data:
            grouped_results[item['file_info']['scan_root']].append(item)
        
        page_size = landscape(A4) if options['orientation'] == 'landscape' else A4
        doc = SimpleDocTemplate(path, pagesize=page_size, topMargin=1.5*cm, bottomMargin=1.5*cm, leftMargin=1.5*cm, rightMargin=1.5*cm)
        styles = getSampleStyleSheet()
        story = []
        num_columns = options['columns']
        col_width = (doc.width / num_columns) - (cm * 0.2 * (num_columns - 1))
        
        for folder, files_with_pages in sorted(grouped_results.items()):
            story.append(Paragraph(self._get_display_path(files_with_pages[0]['file_info']), styles['h2']))
            grid_data = []; row = []
            
            for item_container in files_with_pages:
                item_data = item_container['file_info']
                selected_pages = item_container['selected_pages']
                
                for page_num in selected_pages:
                    cell_content = []
                    try:
                        if item_data['type'] in ('PDF', 'AI'):
                            with fitz.open(item_data['full_path']) as doc_pdf:
                                pix = doc_pdf.load_page(page_num).get_pixmap(dpi=150); img_data = pix.tobytes("png")
                        else:
                            with Image.open(item_data['full_path']) as img:
                                img.thumbnail((400, 400)); img_buffer = io.BytesIO(); img.save(img_buffer, format='PNG'); img_data = img_buffer.getvalue()
                        cell_content.append(ReportLabImage(io.BytesIO(img_data), width=col_width*0.9, height=col_width*0.9, kind='proportional'))
                    except Exception as e: print(f"Impossibile creare anteprima PDF: {e}")
                    page_info = f" (Pag. {page_num + 1}/{item_data['page_count']})" if item_data['page_count'] > 1 else ""
                    cell_content.append(Paragraph(item_data['filename'] + page_info, styles['Normal']))
                    row.append(cell_content)
                    if len(row) == num_columns: grid_data.append(row); row = []
            
            if row: row.extend([""] * (num_columns - len(row))); grid_data.append(row)
            if grid_data:
                table = Table(grid_data, colWidths=[col_width] * num_columns); table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (0,0), (-1,-1), 'CENTER')])); story.append(table)
            story.append(PageBreak())
        
        if story and isinstance(story[-1], PageBreak): story.pop()
        doc.build(story)
        return f"PDF esportato con successo in:\n{path}"

    @Slot(object, bool)
    def on_export_finished(self, message, success):
        self.status_label.setText("Operazione completata.")
        msg_box = QMessageBox.information if success else QMessageBox.critical
        msg_box(self, "Esportazione Completata" if success else "Errore", str(message))

    @Slot(object, bool)
    def on_html_preview_finished(self, message, success):
        if success:
            self.status_label.setText(str(message))
        else:
            self.status_label.setText("Errore durante la creazione dell'anteprima HTML.")
            QMessageBox.critical(self, "Errore Anteprima", str(message))

    def _get_file_details(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            details = {"filename": os.path.basename(file_path), "type": ext.replace('.', '').upper(), "path": os.path.dirname(file_path), "full_path": file_path}
            if ext in ('.jpg', '.jpeg', '.tif', '.tiff'):
                with Image.open(file_path) as img:
                    w, h = img.size; dpi_x, dpi_y = img.info.get('dpi', (DEFAULT_DPI, DEFAULT_DPI))
                    dpi_x = DEFAULT_DPI if dpi_x == 0 else dpi_x; dpi_y = DEFAULT_DPI if dpi_y == 0 else dpi_y
                    w_cm, h_cm = (w/dpi_x)*2.54, (h/dpi_y)*2.54
                    details.update({"page_count": 1, "pages_details": [{"dimensions_cm": f"{w_cm:.2f} x {h_cm:.2f}", "width_cm": w_cm, "height_cm": h_cm}], "dpi_str": f"{int(dpi_x)} DPI"})
                    return details
            elif ext in ('.pdf', '.ai'):
                with fitz.open(file_path) as doc:
                    if len(doc) == 0: return None
                    pages = [{"dimensions_cm": f"{(p.rect.width/72)*2.54:.2f} x {(p.rect.height/72)*2.54:.2f}", "width_cm": (p.rect.width/72)*2.54, "height_cm": (p.rect.height/72)*2.54} for p in doc]
                    details.update({"type": "AI" if ext == '.ai' else "PDF", "page_count": doc.page_count, "pages_details": pages})
                    return details
        except Exception as e:
            print(f"Errore analisi file {file_path}: {e}")
            return {"filename": os.path.basename(file_path), "type": "ERRORE", "path": os.path.dirname(file_path), "full_path": file_path, "page_count": 1, "pages_details": [{"dimensions_cm": "Errore", "width_cm": 0, "height_cm": 0}]}

    def _generate_html_content(self, results, is_print=False):
        grouped_results = defaultdict(list)
        for item in results:
            grouped_results[item['file_info']['scan_root']].append(item)
        
        file_colors = ['#DB4437', '#4285F4', '#F4B400', '#0F9D58', '#AB47BC']
        body_onload = 'onload="window.print()"' if is_print else ''
        css = """<style id="page-orientation-style"></style><style>@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');:root{--list-thumbnail-size:80px;--grid-item-width:200px}body{font-family:'Roboto',sans-serif;margin:0;background-color:#f4f4f4}.controls{display:flex;flex-wrap:wrap;align-items:center;gap:15px;margin-bottom:20px;background:#fff;padding:10px 15px;border-radius:8px;box-shadow:0 2px 4px #0000001a;position:sticky;top:10px;z-index:1000}#printable-content{margin:20px auto;padding:15mm;background:#fff;box-shadow:0 0 10px #0000001a;box-sizing:border-box}.controls button,.controls select{padding:8px 12px;font-size:14px;background-color:#e9e9e9;color:#333;border:1px solid #ccc;border-radius:5px;cursor:pointer;font-weight:700}#search-box{padding:8px;border:1px solid #ccc;border-radius:5px;width:200px}.print-button{background-color:#4285f4;color:#fff;border-color:#4285f4}.control-group{display:flex;align-items:center;gap:5px}.view-switcher button.active{background-color:#4285f4;color:#fff;border-color:#4285f4}.control-group button.size-btn{width:35px;height:35px;font-size:18px;line-height:1}.folder-container{margin-bottom:20px}.folder-header{background-color:#e0e0e0;padding:10px;border-left:5px solid #4285f4;overflow:hidden}.folder-path{font-weight:700;font-size:1.2em;float:left}.folder-stats{float:right;font-size:.9em;color:#555;line-height:1.5em}.grid-container{display:flex;flex-wrap:wrap;justify-content:center;gap:15px;margin-top:15px}.grid-container .item{width:var(--grid-item-width);display:flex;flex-direction:column;border:1px solid #ddd;border-radius:5px;padding:10px;text-align:center;background-color:#fff;page-break-inside:avoid;position:relative;overflow:hidden}.grid-container .item-img{max-width:100%;height:auto;border-radius:3px;object-fit:contain}.grid-container .item-info{text-align:center;flex-grow:1}.list-container{display:flex;flex-direction:column;gap:8px;margin-top:15px}.list-container .item{display:flex;align-items:center;border:1px solid #ddd;border-radius:5px;padding:8px;background-color:#fff;page-break-inside:avoid;position:relative}.list-container .item-img{width:var(--list-thumbnail-size);height:var(--list-thumbnail-size);object-fit:contain;border-radius:3px;margin-right:15px;flex-shrink:0}.list-container .item-info{flex-grow:1;text-align:left}.filename{font-size:.9em;font-weight:700;margin-top:5px;word-wrap:break-word}.dimensions{font-size:.8em;color:#777}.page-indicator{position:absolute;top:5px;right:5px;font-size:.7em;color:#fff;padding:2px 5px;border-radius:3px}.annotation-area{width:100%;box-sizing:border-box;margin-top:8px;padding:5px;border:1px dashed #ccc;border-radius:4px;font-family:sans-serif;resize:vertical;min-height:40px;font-size:14px;font-weight:700;color:#d32f2f}.folder-annotation{margin:10px 0;min-height:50px}.list-container .annotation-area{margin-left:15px}@media print{html,body{width:100%;height:100%;margin:0;padding:0}.controls{display:none}body{background-color:#fff}#printable-content{width:100%;margin:0;padding:0;box-shadow:none}.folder-header{background-color:#f0f0f0!important;-webkit-print-color-adjust:exact}.page-indicator{background-color:var(--bg-color)!important;-webkit-print-color-adjust:exact}.annotation-area{border:1px solid #eee;resize:none;background-color:#fdfdfd!important;-webkit-print-color-adjust:exact}.hide-on-print{display:none!important}}</style>"""
        js_script = """<script>let state={view:"grid",gridItemWidth:200,listThumbSize:80};function switchView(e){if(state.view===e)return;state.view=e,document.querySelectorAll(".content-container").forEach(t=>{t.classList.remove("grid-container","list-container"),t.classList.add(e+"-container")}),document.getElementById("btn-grid").classList.toggle("active","grid"===e),document.getElementById("btn-list").classList.toggle("active","list"===e)}function changeSize(e){"grid"===state.view?(state.gridItemWidth=Math.max(80,Math.min(600,state.gridItemWidth+40*e)),document.documentElement.style.setProperty("--grid-item-width",state.gridItemWidth+"px")):(state.listThumbSize=Math.max(40,Math.min(200,state.listThumbSize+20*e)),document.documentElement.style.setProperty("--list-thumbnail-size",state.listThumbSize+"px"))}function updatePrintStyle(){let e=document.querySelector('input[name="orientation"]:checked').value;document.getElementById("page-orientation-style").innerHTML=`@page { size: A4 ${e}; margin: 1.5cm; }`;let t=document.getElementById("printable-content");t.style.width="portrait"===e?"180mm":"267mm"}function filterFiles(){let e=document.getElementById("search-box").value.toLowerCase();document.querySelectorAll(".folder-container").forEach(t=>{let i=t.querySelectorAll(".item"),l=0,n=0;i.forEach(t=>{let i=t.querySelector(".filename").textContent.toLowerCase();i.includes(e)?(t.style.display="flex",l++,n+=parseFloat(t.dataset.area)):t.style.display="none"});let a=t.querySelector(".folder-header"),s=t.querySelector(".folder-stats"),d=a.dataset.originalFiles,o=a.dataset.originalPages,r=a.dataset.originalSqm;""===e.trim()?s.textContent=`File: ${d} | Pagine: ${o} | Area: ${r} m²`:s.textContent=`File: ${l} (di ${d}) | Area: ${n.toFixed(2)} m²`,t.style.display=l>0?"":"none"})}function prepareAndPrint(){document.querySelectorAll(".annotation-area").forEach(e=>{e.classList.toggle("hide-on-print",""===e.value.trim())}),window.print()}document.addEventListener("DOMContentLoaded",()=>{switchView("grid"),updatePrintStyle(),document.documentElement.style.setProperty("--grid-item-width",state.gridItemWidth+"px"),document.documentElement.style.setProperty("--list-thumbnail-size",state.listThumbSize+"px")});</script>"""
        body = f"""<body {body_onload}><div class="controls"><button onclick="prepareAndPrint()" class="print-button">Stampa Pagina</button><input type="text" id="search-box" onkeyup="filterFiles()" placeholder="Cerca per nome file..."><div class="control-group"><label>Orientamento:</label><input type="radio" id="portrait" name="orientation" value="portrait" checked onchange="updatePrintStyle()"><label for="portrait">Verticale</label><input type="radio" id="landscape" name="orientation" value="landscape" onchange="updatePrintStyle()"><label for="landscape">Orizzontale</label></div><div class="view-switcher control-group"><button id="btn-grid" onclick="switchView('grid')">Griglia</button><button id="btn-list" onclick="switchView('list')">Elenco</button></div><div class="size-controls control-group"><label>Dimensione:</label><button class="size-btn" onclick="changeSize(-1)">-</button><button class="size-btn" onclick="changeSize(1)">+</button></div></div><div id="printable-content">"""
        
        for folder, files_with_pages in sorted(grouped_results.items()):
            display_folder = self._get_display_path(files_with_pages[0]['file_info'])
            num_files = len(files_with_pages)
            total_pages = sum(len(f['selected_pages']) for f in files_with_pages)
            total_sqm = sum(f['file_info']['pages_details'][p_idx]['width_cm'] * f['file_info']['pages_details'][p_idx]['height_cm'] for f in files_with_pages for p_idx in f['selected_pages']) / 10000
            
            body += f"""<div class="folder-container"><div class="folder-header" data-original-files="{num_files}" data-original-pages="{total_pages}" data-original-sqm="{total_sqm:.2f}"><span class="folder-stats">File: {num_files} | Pagine: {total_pages} | Area: {total_sqm:.2f} m²</span><span class="folder-path">{display_folder}</span></div><textarea class="annotation-area folder-annotation" placeholder="Aggiungi un'annotazione per questa cartella..."></textarea><div class="content-container grid-container">"""
            
            for file_index, item_container in enumerate(files_with_pages):
                item_data = item_container['file_info']
                selected_pages = item_container['selected_pages']
                page_count = item_data.get('page_count', 1)
                current_color = file_colors[file_index % len(file_colors)]
                
                for page_num in selected_pages:
                    try:
                        page_details = item_data["pages_details"][page_num]
                        area_sqm = (page_details.get('width_cm', 0) * page_details.get('height_cm', 0)) / 10000
                        if item_data['type'] in ('PDF', 'AI'):
                            with fitz.open(item_data['full_path']) as doc_pdf:
                                pix = doc_pdf.load_page(page_num).get_pixmap(dpi=150); img_data = pix.tobytes("png")
                        else:
                            with Image.open(item_data['full_path']) as img:
                                img.thumbnail((400, 400)); img_buffer = io.BytesIO(); img.save(img_buffer, format='PNG'); img_data = img_buffer.getvalue()
                        b64_img = base64.b64encode(img_data).decode('utf-8')
                        img_src = f"data:image/png;base64,{b64_img}"
                        dpi_info = f"({item_data['dpi_str']})" if item_data.get('dpi_str') else ""
                        body += f'<div class="item" data-area="{area_sqm}">'
                        if page_count > 1: body += f'<div class="page-indicator" style="background-color: {current_color}; --bg-color: {current_color};">Pag. {page_num + 1}/{page_count}</div>'
                        body += f'''<img class="item-img" src="{img_src}" alt="Anteprima"><div class="item-info"><div class="filename">{html.escape(item_data["filename"])}</div><div class="dimensions">{html.escape(page_details["dimensions_cm"])} cm {html.escape(dpi_info)}</div></div><textarea class="annotation-area" placeholder="Annotazione..."></textarea></div>'''
                    except Exception as e: print(f"Impossibile creare anteprima HTML: {e}")
            body += '</div></div>'
        body += '</div></body>'
        return f"<!DOCTYPE html><html><head><meta charset='UTF-8'><title>Anteprima Report</title>{css}{js_script}</head>{body}</html>"

    def _get_display_path(self, file_info):
        norm_path = os.path.normpath(file_info['path'])
        scan_root = os.path.normpath(file_info.get('scan_root', ''))
        if not scan_root: return os.path.basename(norm_path)
        try:
            if os.path.splitdrive(norm_path)[0].upper() == os.path.splitdrive(scan_root)[0].upper():
                rel_path = os.path.relpath(norm_path, scan_root)
                return os.path.join(os.path.basename(scan_root), rel_path) if rel_path != '.' else os.path.basename(scan_root)
            return os.path.basename(norm_path)
        except ValueError:
            return os.path.basename(norm_path)


    def update_theme(self):
        pass

def create_tab(tab_widget, winfile_app):
    tab_name = "Liste e anteprime"
    app_instance = FileScannerApp(master=tab_widget, winfile_app=winfile_app)
    tab_widget.addTab(app_instance, tab_name)
    return tab_name, app_instance
