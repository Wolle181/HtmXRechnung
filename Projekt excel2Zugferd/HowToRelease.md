# HowToRelease – Excel2ZUGFeRD

## Voraussetzungen

Folgende Dateien müssen im Projektverzeichnis vorliegen, bevor ein Release gebaut wird:

- `excel2zugferd.exe` (PyInstaller-Build)
- `Excel2Zugferd.xlam` (Excel Add-in, erzeugt durch `Create-Excel2Zugferd.ps1`)
- `_internal\` (PyInstaller-Laufzeitdateien)
- `Install-Excel2Zugferd.ps1` / `Install.bat`
- `Uninstall-Excel2Zugferd.ps1` / `Uninstall.bat`

Das XLAM neu bauen (falls nötig):
```powershell
powershell -ExecutionPolicy Bypass -File Create-Excel2Zugferd.ps1
```

---

## Release bauen

```powershell
powershell -ExecutionPolicy Bypass -File Create-Release.ps1
```

Das Script legt folgende Struktur an:

```
Release\
├── Excel2ZugferdSetup.bat          ← Endanwender: Doppelklick
├── Excel2ZugferdSetup.ps1
└── Excel2ZugferdSetupPayload\
    ├── excel2zugferd.exe
    ├── Excel2Zugferd.xlam
    ├── Install-Excel2Zugferd.ps1
    ├── Install.bat
    ├── Uninstall-Excel2Zugferd.ps1
    ├── Uninstall.bat
    └── _internal\
```

---

## Weitergabe

Den gesamten Ordner `Release\` zippen und weitergeben.  
Empfänger entpacken das Archiv und starten `Excel2ZugferdSetup.bat` per Doppelklick.

---

## Was das Setup beim Endanwender tut

1. Erstellt `C:\Rechnungen\Excel2Zugferd\`
2. Kopiert alle Programmdateien dorthin
3. Startet `Install.bat`, das das Excel-AddIn für den aktuellen Benutzer registriert

Nach dem nächsten Excel-Start erscheint der Tab **Excel2ZUGFeRD** im Ribbon.

---

## Deinstallation (beim Endanwender)

```
C:\Rechnungen\Excel2Zugferd\Uninstall.bat
```

Das Script entfernt:
- Den Registry-Eintrag `OPEN`/`OPENx` (Auto-Laden beim Excel-Start)
- Den Eintrag aus dem **Add-in Manager** (sichtbare Liste in Excel-Optionen)
- Den **AddInLoadTimes**-Cache
- Die Datei `%APPDATA%\Microsoft\AddIns\Excel2Zugferd.xlam`

Nach einem Excel-Neustart ist kein Eintrag mehr sichtbar.  
Die Programmdateien unter `C:\Rechnungen\Excel2Zugferd\` müssen anschließend manuell gelöscht werden.
