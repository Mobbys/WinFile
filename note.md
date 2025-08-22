GITHUB_TOKEN = "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" 
# Prima di compilare bisogna cambiare la versione sia in winfile.py che in setup.py mettendo la stessa
Comando per compilare   python setup.py build

Le librerie principali da installare per l'applicazione WinFile sono:

customtkinter: per la creazione dell'interfaccia grafica moderna.

Pillow (PIL): per la manipolazione delle immagini.

PyMuPDF (la libreria che corrisponde a fitz nel codice): per la gestione e l'analisi dei file PDF.

reportlab: per la creazione di report professionali in formato PDF.

TkinterDnD2: per la funzionalità di trascinamento e rilascio (Drag & Drop).

requests: per la gestione delle richieste HTTP, usata per il controllo degli aggiornamenti.

packaging: per confrontare le versioni del software.

pyinstaller (opzionale): per creare un eseguibile dell'applicazione.

ctypes e wintypes (incorporati in Python): per la gestione degli appunti formattati su Windows.
