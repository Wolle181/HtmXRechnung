● Fertig. Die Struktur, die das Script erzeugt:

  Release\
  ├── Excel2ZugferdSetup.bat        ← Endanwender: Doppelklick zum Installieren
  ├── Excel2ZugferdSetup.ps1        ← eigentliche Setup-Logik
  └── Excel2ZugferdSetupPayload\
      ├── excel2zugferd.exe
      ├── Excel2Zugferd.xlam
      ├── Install-Excel2Zugferd.ps1
      ├── Install.bat
      ├── Uninstall-Excel2Zugferd.ps1
      ├── Uninstall.bat
      └── _internal\

  Aufruf (einmalig beim Bauen):
  powershell -ExecutionPolicy Bypass -File Create-Release.ps1

  Was das Excel2ZugferdSetup.bat beim Endanwender macht:
  1. Erstellt C:\Rechnungen\Excel2Zugferd\
  2. Kopiert alles aus dem Payload dorthin
  3. Startet Install.bat aus dem Installationsverzeichnis (das wiederum das Excel-AddIn registriert)

  Zum Weitergeben reicht es, den ganzen Release\-Ordner zu zippen.
  