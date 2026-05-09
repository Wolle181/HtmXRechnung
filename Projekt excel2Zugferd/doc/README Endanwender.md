# README zur Installation der Excel2Zugferd-Lösung inkl. Excel-Addin

## Voraussetzungen

Die als ZIP-Datei erfolgte Auslieferung ist entpackt worden in einem beliebigen, temporären Verzeichnis.

## Installation

Als Endanwender starten Sie in dem Ordner, in den Sie das Auslieferungs-ZIP entpackt haben, die Datei `Excel2ZugferdSetup.bat` per Doppelklick.

---

### Was das Setup tut

1. Erstellt das Verzeichnis `C:\Rechnungen\Excel2Zugferd\`
2. Kopiert alle Programmdateien dorthin
3. Startet `Install.bat`, das das Excel-AddIn für den aktuellen Benutzer registriert

Nach dem nächsten Excel-Start erscheint der Tab **Excel2ZUGFeRD** im Ribbon.

Das temporäre Verzeichnis wird nach diesem Schritt nicht mehr benötigt und kann gelöscht werden.

---

## Deinstallation des Excel-Addins

Falls das excel2Zugferd-AddIn vollständig entfernt werden soll, muss das folgende Batch aufgerufen werden:

```cmd
C:\Rechnungen\Excel2Zugferd\Uninstall.bat
```


Nach dem nächsten Excel-Start ist das Ribbon-Tab **Excel2ZUGFeRD** nicht mehr vorhanden.

### Hinweis

Die Programmdateien unter `C:\Rechnungen\Excel2Zugferd\` müssen anschließend manuell gelöscht werden. Dies ist beabsichtigt.
