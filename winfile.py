# winfile.py - v3.2 (PySide6) - Layout a piena altezza
import sys
import importlib.util
from pathlib import Path
import json
import logging
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout, 
                               QVBoxLayout, QLabel, QComboBox, QTabWidget, QFrame)
from PySide6.QtCore import Qt, QObject, QEvent
from PySide6.QtGui import QFont, QCursor

# --- IMPOSTAZIONE DEL LOGGING ---
def setup_logging():
    try:
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        log_file_path = os.path.join(desktop_path, 'winfile_log.txt')
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file_path, mode='w'),
                logging.StreamHandler()
            ]
        )
        logging.info("Logging avviato.")
    except Exception as e:
        print(f"Errore durante l'impostazione del logging: {e}")

CONFIG_FILE = "config.json"

def save_config(settings):
    with open(CONFIG_FILE, "w") as f:
        json.dump(settings, f, indent=4)

def load_config():
    if not Path(CONFIG_FILE).exists():
        return {"theme": "Dark"}
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        return {"theme": "Dark"}

# --- STILI PER L'INTERFACCIA (QSS) ---
STYLESHEET_DARK = """
    QWidget {
        background-color: #2B2B2B;
        color: #FFFFFF;
        font-family: Segoe UI;
        font-size: 10pt;
    }
    QMainWindow {
        background-color: #2B2B2B;
    }
    QTabWidget::pane {
        border: 1px solid #343638;
        border-top: none; /* Rimuove il bordo superiore del pannello delle schede */
        border-radius: 0px;
    }
    QTabBar::tab {
        background: #2B2B2B;
        border: 1px solid #343638;
        padding: 8px 15px;
        border-top-left-radius: 4px;
        border-top-right-radius: 4px;
        border-bottom: none; /* Rimuove il bordo inferiore della scheda */
    }
    QTabBar::tab:selected {
        background: #1f6aa5;
        border-bottom-color: #1f6aa5;
    }
    QTabBar::tab:!selected:hover {
        background: #343638;
    }
    QPushButton {
        background-color: #1f6aa5;
        border-radius: 4px;
        padding: 6px 12px;
        border: 1px solid #1f6aa5;
    }
    QPushButton:hover {
        background-color: #3484d0;
    }
    QPushButton:pressed {
        background-color: #1a5a90;
    }
    QComboBox {
        border: 1px solid #343638;
        border-radius: 4px;
        padding: 4px;
        background-color: #343638;
        min-width: 100px; /* Larghezza minima per il combobox */
    }
    QComboBox::drop-down {
        border: none;
    }
    QHeaderView::section {
        background-color: #343638;
        padding: 4px;
        border: 1px solid #2B2B2B;
    }
    QTreeWidget {
        border: 1px solid #343638;
        border-radius: 4px;
    }
    #BottomBar {
        border-top: 1px solid #343638;
    }
"""

STYLESHEET_LIGHT = """
    QWidget {
        background-color: #FFFFFF;
        color: #000000;
        font-family: Segoe UI;
        font-size: 10pt;
    }
    QMainWindow {
        background-color: #F0F0F0;
    }
    QTabWidget::pane {
        border: 1px solid #EAEAEA;
        border-top: none; /* Rimuove il bordo superiore del pannello delle schede */
        border-radius: 0px;
    }
    QTabBar::tab {
        background: #F0F0F0;
        border: 1px solid #EAEAEA;
        padding: 8px 15px;
        border-top-left-radius: 4px;
        border-top-right-radius: 4px;
        border-bottom: none; /* Rimuove il bordo inferiore della scheda */
    }
    QTabBar::tab:selected {
        background: #3484d0;
        color: white;
        border-bottom-color: #3484d0;
    }
    QTabBar::tab:!selected:hover {
        background: #EAEAEA;
    }
    QPushButton {
        background-color: #3484d0;
        color: white;
        border-radius: 4px;
        padding: 6px 12px;
        border: 1px solid #3484d0;
    }
    QPushButton:hover {
        background-color: #4997e0;
    }
    QPushButton:pressed {
        background-color: #2a70b4;
    }
    QComboBox {
        border: 1px solid #CCCCCC;
        border-radius: 4px;
        padding: 4px;
        background-color: #F0F0F0;
        min-width: 100px; /* Larghezza minima per il combobox */
    }
    QComboBox::drop-down {
        border: none;
    }
    QHeaderView::section {
        background-color: #EAEAEA;
        padding: 4px;
        border: 1px solid #FFFFFF;
    }
    QTreeWidget {
        border: 1px solid #EAEAEA;
        border-radius: 4px;
    }
    #BottomBar {
        background-color: #F0F0F0;
        border-top: 1px solid #CCCCCC;
    }
"""

class DragDropEventFilter(QObject):
    def __init__(self, window):
        super().__init__()
        self.window = window

    def eventFilter(self, watched, event):
        if event.type() not in [QEvent.Type.DragEnter, QEvent.Type.DragMove, QEvent.Type.Drop]:
            return super().eventFilter(watched, event)

        tab_widget = self.window.tab_view
        pos_in_tab_widget = tab_widget.mapFromGlobal(QCursor.pos())
        
        if not tab_widget.rect().contains(pos_in_tab_widget):
            return super().eventFilter(watched, event)

        current_tab = tab_widget.currentWidget()
        if not hasattr(current_tab, 'handle_drop_event'):
            return super().eventFilter(watched, event)

        if event.type() in [QEvent.Type.DragEnter, QEvent.Type.DragMove]:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True

        elif event.type() == QEvent.Type.Drop:
            if event.mimeData().hasUrls():
                current_tab.handle_drop_event(event)
                event.acceptProposedAction()
                return True

        return super().eventFilter(watched, event)

class WinFileApp(QMainWindow):
    def __init__(self):
        super().__init__()
        logging.info("Inizializzazione di WinFileApp con PySide6.")
        
        self.app_instances = {}
        self.setWindowTitle("WinFile v3.2 (PySide6)")
        self.setGeometry(100, 100, 1200, 700)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # --- MODIFICA LAYOUT PRINCIPALE ---
        # Il layout principale ora è verticale per contenere le schede e la barra inferiore
        self.main_layout = QVBoxLayout(self.central_widget)
        # Rimuove margini e spaziatura per un layout a piena altezza
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # L'area delle schede occupa tutto lo spazio disponibile
        self.tab_view = QTabWidget()
        self.main_layout.addWidget(self.tab_view, 1) # Il '1' indica che si espande

        # --- NUOVA BARRA INFERIORE ---
        self.bottom_bar = QFrame()
        self.bottom_bar.setObjectName("BottomBar") # ID per lo stile QSS
        self.bottom_bar.setFixedHeight(40)
        self.bottom_bar_layout = QHBoxLayout(self.bottom_bar)
        self.bottom_bar_layout.setContentsMargins(10, 0, 10, 0)

        # Aggiunta del selettore del tema alla barra inferiore
        self.theme_label = QLabel("Tema:")
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        self.theme_combo.currentTextChanged.connect(self.change_theme_event)
        
        self.bottom_bar_layout.addWidget(self.theme_label)
        self.bottom_bar_layout.addWidget(self.theme_combo)
        self.bottom_bar_layout.addStretch() # Spinge i widget a sinistra

        # Aggiunta della barra inferiore al layout principale
        self.main_layout.addWidget(self.bottom_bar)
        
        # --- FINE MODIFICHE LAYOUT ---

        self.load_apps()
        
        self.config = load_config()
        self.theme_combo.setCurrentText(self.config.get("theme", "Dark"))
        self.change_theme_event(self.config.get("theme", "Dark"))
        logging.info("WinFileApp inizializzata con successo.")

    def change_theme_event(self, theme: str):
        logging.info(f"Cambiando il tema in: {theme}")
        if theme == "Dark": self.setStyleSheet(STYLESHEET_DARK)
        else: self.setStyleSheet(STYLESHEET_LIGHT)
        self.config["theme"] = theme
        save_config(self.config)
        self.broadcast_theme_update()

    def load_apps(self):
        logging.info("Inizio caricamento delle app...")
        apps_path = Path("apps")
        if not apps_path.is_dir():
            logging.error("La cartella 'apps' non è stata trovata.")
            return

        for file_path in sorted(apps_path.glob("app_*.py")):
            module_name = file_path.stem
            try:
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                app_module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(app_module)
                if hasattr(app_module, "create_tab"):
                    try:
                        tab_name, instance = app_module.create_tab(self.tab_view, self)
                        if tab_name and instance: self.app_instances[tab_name] = instance
                        logging.info(f"Modulo '{module_name}' caricato con successo nella scheda '{tab_name}'.")
                    except Exception as e:
                        logging.exception(f"Errore durante la creazione della scheda per '{module_name}'.")
                        error_label = QLabel(f"Errore nel caricamento di:\n{module_name}\n\n{e}")
                        error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        self.tab_view.addTab(error_label, f"! {module_name}")
            except Exception as e:
                logging.exception(f"Errore critico durante il caricamento del modulo '{module_name}'.")

    def broadcast_theme_update(self):
        for app in self.app_instances.values():
            if hasattr(app, 'update_theme'):
                try: app.update_theme()
                except Exception: logging.exception(f"Errore durante l'aggiornamento del tema.")


if __name__ == "__main__":
    setup_logging()
    app = QApplication(sys.argv)
    window = WinFileApp()
    
    event_filter = DragDropEventFilter(window)
    app.installEventFilter(event_filter)
    
    window.show()
    sys.exit(app.exec())
