# WinFile

## Struttura Generale
**WinFile** è un'applicazione desktop modulare basata su Python e `customtkinter`. Funge da launcher e contenitore per diverse utilità ("App") dedicate alla gestione, analisi e manipolazione di file grafici (immagini, PDF) e documenti tecnici.

### Componenti Principali
*   **Main Application (`winfile.py`)**:
    *   Gestisce la finestra principale e il caricamento dinamico dei moduli dalla cartella `apps/`.
    *   Sistema di aggiornamento automatico tramite GitHub Release.
    *   Gestione delle impostazioni globali (tema, dimensioni finestra).
    *   Supporto Drag & Drop globale.

## Applicazioni (Moduli)

### 1. Controllo Immagini (`apps/app_controllo_immagini.py`)
Strumento per l'analisi tecnica e la modifica rapida di file immagine.
*   **Funzioni Principali**:
    *   Visualizzazione dettagli tecnici: Dimensioni (px/cm), Risoluzione (DPI), Metodo Colore, Profilo Colore.
    *   **Calcolatrice DPI/Dimensioni**: Ricalcola automaticamente i valori collegati modificando uno dei parametri (es. cambia DPI -> aggiorna cm).
    *   **Trasformazioni**: Ruota, Specchia (Orizzontale/Verticale).
    *   **Ritaglio (Crop)**: Strumento di ritaglio interattivo con possibilità di inserire coordinate manuali.
    *   **Lente d'ingrandimento**: Zoom locale per ispezione dettagli.
    *   **Salva con nome**: Esportazione dell'immagine elaborata.

### 2. Controllo PDF (`apps/app_controllo_pdf.py`)
Editor leggero per file PDF e AI (Adobe Illustrator).
*   **Funzioni Principali**:
    *   **Visualizzazione**: Anteprima pagine con navigazione e miniature.
    *   **Selezione Pagine**: Interfaccia visuale per selezionare pagine specifiche.
    *   **Modifica Pagine**:
        *   **Aggiungi Margini**: Aggiunge spazio bianco o colorato attorno alle pagine.
        *   **Ridimensiona**: Cambia il formato pagina (es. A4, A3) mantenendo le proporzioni o adattando il contenuto.
        *   **Ritaglia**: Crop delle pagine.
        *   **Elimina**: Rimozione pagine indesiderate.
    *   **Unione/Inserimento**: Drag & Drop di file esterni per aggiungere o sostituire pagine/file.

### 3. Liste Anteprime (`apps/app_liste_anteprime.py`)
Generatore di report e cataloghi visivi per file in cartelle e sottocartelle.
*   **Funzioni Principali**:
    *   **Scansione Ricorsiva**: Analizza cartelle per trovare Immagini e PDF.
    *   **Tabella Dati**: Mostra nome file, dimensioni, area (mq) e percorso.
    *   **Ordinamento**: Sort per nome, dimensione, area, percorso.
    *   **Esportazione**:
        *   **Anteprima HTML**: Genera una galleria web con miniature e annotazioni.
        *   **PDF**: Crea un documento PDF con la lista o le miniature.
        *   **CSV**: Esporta i dati tabellari per Excel.
        *   **Stampa**: Stampa diretta della lista file.
    *   **Opzioni di Layout**: Configurazione colonne, orientamento pagina, dimensione font annotazioni.

### 4. Simulazione Quote (`apps/app_simulazione_quote.py`)
Strumento per misurazioni e simulazioni tecniche su immagini (es. disegni tecnici o foto).
*   **Funzioni Principali**:
    *   **Caricamento Immagini**: Supporto per file grafici/vettoriali (tramite conversione).
    *   **Correzione Prospettica**: Strumento a 4 punti per raddrizzare immagini prospettiche.
    *   **Ritaglio**: Crop dell'area di interesse.
    *   **Strumenti di Misura**: (Dal codice sembra includere funzionalità per definire punti di misura e scale, anche se in fase di sviluppo/affinamento).
    *   **Lente di Precisione**: Zoom dinamico durante il posizionamento dei punti.

## Librerie Esterne Chiave
*   `customtkinter`: Interfaccia grafica moderna.
*   `Pillow (PIL)`: Elaborazione immagini.
*   `PyMuPDF (fitz)`: Gestione e rendering PDF.
*   `reportlab`: Generazione PDF.
*   `tkinterdnd2`: Supporto Drag & Drop avanzato.
