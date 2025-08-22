# WinFile

Questo documento fornisce una panoramica completa dell'applicazione desktop WinFile

## 1. Panoramica Generale

WinFile è un'applicazione modulare, leggera e versatile, progettata per aiutare professionisti e amatori a gestire e analizzare file multimediali, con un'attenzione particolare a immagini (JPEG, TIFF) e documenti (PDF, AI). La sua architettura a schede permette di espandere facilmente le funzionalità.

## 2. Modulo "Liste e anteprime"

Questo modulo è il cuore del sistema di analisi di massa e reportistica.

### Caratteristiche principali:

- **Scansione di file**: Supporta il trascinamento di file e intere cartelle. L'applicazione rileva automaticamente i formati supportati (`.jpg`, `.jpeg`, `.tif`, `.tiff`, `.pdf`, `.ai`)
- **Acquisizione dati**: Estrae metadati cruciali come le dimensioni fisiche in centimetri, il numero di pagine e il percorso del file
- **Visualizzazione gerarchica**: Organizza i documenti multi-pagina (come i PDF) in una struttura ad albero intuitiva, semplificando la navigazione
- **Anteprime**: Visualizza una miniatura dell'immagine o della pagina selezionata nella lista con le relative informazioni

### Funzionalità di Esportazione:

- **Esportazione CSV**: Esporta tutti i dati in un file CSV, ideale per l'analisi in un foglio di calcolo
- **Anteprima HTML**: Genera un report web interattivo con miniature, dettagli e campi per le note
- **Report PDF**: Crea un documento PDF professionale con layout personalizzabile, perfetto per la stampa o la condivisione

### Interazione Utente:

- Un menu contestuale con clic destro offre opzioni rapide per copiare, stampare o rimuovere elementi
- La copia formattata (solo su Windows) permette di incollare tabelle direttamente in applicazioni come Excel e Word

## 3. Modulo: "Controllo immagini"

Questo modulo si concentra sull'analisi dettagliata di singole immagini.

### Caratteristiche principali:

- **Dettagli approfonditi**: Mostra dimensioni in pixel, rapporto d'aspetto, modo colore, dimensione del file e risoluzione (DPI)
- **Conversione interattiva**: I campi di input sono collegati, permettendo di calcolare le dimensioni in cm modificando i DPI, e viceversa
- **Calcolo della distanza minima di visione**: Stima la distanza ottimale per la visualizzazione dell'immagine, fondamentale per la pianificazione di stampe di grandi dimensioni
