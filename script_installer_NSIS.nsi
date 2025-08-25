; ==================================================================
; == Script NSIS per l'installazione di WinFile 3.0
; ==================================================================

; ##################################################################
; ### 1. DEFINIZIONI E IMPOSTAZIONI GENERALI
; ##################################################################

!define NOME_PRODOTTO "WinFile"
!define VERSIONE "3.1"
!define NOME_FILE_EXE "WinFile.exe"
!define CARTELLA_SORGENTE "build\WinFile ${VERSIONE}" ; Cartella con i file da installare
!define FILE_ICONA "icona.ico"
!define NOME_SETUP "WinFile ${VERSIONE} - Installer.exe"

; --- Impostazioni dell'interfaccia (Modern UI 2) ---
; Questo rende l'installer più moderno
!include "MUI2.nsh"

; --- Icone ---
; Icona per il file setup.exe e per il programma di disinstallazione
!define MUI_ICON "${FILE_ICONA}"
!define MUI_UNICON "${FILE_ICONA}"

; --- Proprietà Principali dell'Installer ---
Name "${NOME_PRODOTTO} ${VERSIONE}"
OutFile "${NOME_SETUP}"
InstallDir "$APPDATA\${NOME_PRODOTTO}" ; Installa in una nuova cartella sul Desktop dell'utente
RequestExecutionLevel user ; Non richiede permessi di amministratore, perfetto per il Desktop

; ##################################################################
; ### 2. PAGINE DELL'INTERFACCIA GRAFICA
; ##################################################################

; Definisce le pagine che l'utente vedrà durante l'installazione
!insertmacro MUI_PAGE_DIRECTORY   ; Pagina per confermare/cambiare la cartella (preimpostata sul Desktop)
!insertmacro MUI_PAGE_INSTFILES   ; Pagina che mostra il progresso dell'installazione

; Definisce le pagine per la disinstallazione
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Imposta la lingua dell'interfaccia
!insertmacro MUI_LANGUAGE "Italian"

; ##################################################################
; ### 3. LOGICA DI AGGIORNAMENTO (DISINSTALLAZIONE PRECEDENTE)
; ##################################################################

Function .onInit
  ; Questa funzione viene eseguita prima che appaia qualsiasi finestra.
  ; Cerca se una versione precedente del programma è già installata.

  ; Legge la stringa di disinstallazione della vecchia versione dal registro
  ReadRegStr $R0 HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}" "UninstallString"

  ; Se la stringa non è vuota (cioè, il programma è installato), allora procede
  StrCmp $R0 "" done

  ; Esegue il vecchio uninstaller in modalità silenziosa (/S) e attende la sua conclusione.
  ; Questo rimuove la vecchia versione prima di installare la nuova.
  ExecWait '"$R0" /S _?=$INSTDIR'

done:
FunctionEnd

; ##################################################################
; ### 4. SEZIONE DI INSTALLAZIONE (COSA FA L'INSTALLER)
; ##################################################################

Section "Installazione Principale" SEC_MAIN

  ; Imposta la cartella di destinazione dove verranno copiati i file
  ; $INSTDIR è la variabile che contiene il percorso, es: C:\Users\TuoNome\Desktop\WinFile
  SetOutPath $INSTDIR

  ; Copia TUTTI i file e le sottocartelle dalla cartella sorgente alla destinazione
  ; Il comando /r è ricorsivo, quindi prende tutto il contenuto.
  File /r "${CARTELLA_SORGENTE}\*.*"

  ; Copia esplicitamente l'icona nella cartella di installazione
  File "${FILE_ICONA}"

  ; --- Creazione dei collegamenti ---
  ; Crea un collegamento sul Desktop che punta all'eseguibile nella sua nuova cartella
  CreateShortCut "$DESKTOP\${NOME_PRODOTTO}.lnk" "$INSTDIR\${NOME_FILE_EXE}" "" "$INSTDIR\${FILE_ICONA}" 0

  ; --- Scrittura delle informazioni per la disinstallazione ---
  ; Queste informazioni permettono a Windows di mostrare il programma in "Installazione Applicazioni"
  ; Vengono scritte in HKCU (HKEY_CURRENT_USER) perché non servono permessi di admin.
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}" "DisplayName" "${NOME_PRODOTTO}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}" "DisplayIcon" "$INSTDIR\${FILE_ICONA}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}" "DisplayVersion" "${VERSIONE}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}" "Publisher" "Azienda Interna" ; Puoi cambiarlo se vuoi

  ; Crea il file uninstall.exe nella cartella di installazione
  WriteUninstaller "$INSTDIR\uninstall.exe"

SectionEnd

; ##################################################################
; ### 5. SEZIONE DI DISINSTALLAZIONE
; ##################################################################

Section "Uninstall"

  ; Rimuove l'intera cartella del programma dal Desktop
  ; RMDir /r è ricorsivo e rimuove cartella, sottocartelle e tutti i file.
  RMDir /r "$INSTDIR"

  ; Rimuove il collegamento creato sul Desktop
  Delete "$DESKTOP\${NOME_PRODOTTO}.lnk"

  ; Rimuove le chiavi di registro per pulire la lista dei programmi installati
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${NOME_PRODOTTO}"

SectionEnd
