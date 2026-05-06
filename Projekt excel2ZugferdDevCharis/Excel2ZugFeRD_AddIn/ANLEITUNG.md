# Excel Add-In für ZUGFeRD PDF Erstellung

## Anleitung zum Erstellen des Add-Ins

Das Add-In ist eine `.xlam` Datei (Excel Macro-Enabled Add-In), die einen Button zum Erstellen von ZUGFeRD PDFs bereitstellt.

### Schritt 1: Neues Add-In erstellen

1. **Excel öffnen** → **Neue leere Arbeitsmappe**
2. **Datei** → **Speichern unter** 
   - Dateityp: **Excel-Add-In (*.xlam)**
   - Name: `Excel2ZugFeRD_AddIn`
   - Speicherort: `C:\Users\BENUTZERNAME\AppData\Roaming\Microsoft\AddIns\`

### Schritt 2: VBA-Code einfügen

1. **Alt + F11** → Visual Basic Editor öffnen
2. **Rechtsklick auf "VBAProject"** → **Module einfügen**
3. **Module_ZugFeRD.bas in den Editor kopieren** (siehe Schritt 5)

### Schritt 3: Ribbon UI konfigurieren

Die Datei muss als `.xlam` mit Custom Ribbon exportiert werden:

1. **Speichern & Schließen** (Arbeitsmappe)
2. **Die .xlam Datei als ZIP umbenennen** (`.xlam.zip`)
3. **Entpacken** und folgende Struktur erstellen:

```
Excel2ZugFeRD_AddIn.xlam.zip/
├── xl/
│   ├── workbook.xml
│   ├── workbook.xml.rels
│   ├── customUI/
│   │   └── customUI.xml (siehe Schritt 6)
│   └── ...
├── _rels/
├── [Content_Types].xml
└── ...
```

4. **customUI.xml Version eintragen** in `_rels/.rels`:
```xml
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/ribbon/ui" Target="customUI/customUI.xml"/>
```

5. **Zurück in ZIP umbenennen** auf `.xlam`

### Schritt 4: Im Speicher schnell erstellen (EMPFOHLEN)

Alternativ: **Direkt Dateien vorbereiten:**

- `customUI.xml` → [See customUI.xml](customUI.xml)
- `Module_ZugFeRD.bas` → [See Module_ZugFeRD.bas](Module_ZugFeRD.bas)

### Schritt 5: Einfach mit VBA + Button

**NOCH EINFACHER - Ohne Ribbon XML:**

1. Neue Arbeitsmappe als `.xlam` speichern
2. VBA-Code einfügen
3. Ein Button auf dem Sheet hinzufügen (FormControl)
4. Button-Macro verknüpfen: `CreateZugFeRDPDF`

**Das ist die schnellste Lösung!**

### Schritt 6: Add-In installieren

1. `.xlam` in folgendes Verzeichnis kopieren:
   ```
   C:\Users\[BENUTZERNAME]\AppData\Roaming\Microsoft\AddIns\
   ```

2. **Excel öffnen** → **Datei** → **Optionen** → **Add-Ins**
3. **Verwalten: Excel-Add-Ins** → **Durchsuchen...**
4. `Excel2ZugFeRD_AddIn.xlam` auswählen
5. **OK**

### Schritt 7: Testen

1. **Neue Arbeitsmappe öffnen**
2. **Startseite** → Neuer Button sollte sichtbar sein
3. Button klicken → ZUGFeRD PDF wird erstellt

## Icon Vorschlag

Das beste Icon würde sein:
- **PDF-Symbol** (rotes "PDF" Dokument)
- **+** oder Zahnrad zum Kennzeichnen der Aktion
- **Farben:** Rot für PDF, Blau für Zahlung/Geschäft

**Icon-Beschreibung als PNG (128x128px):**
```
┌─────────────────┐
│  📄 PDF Zone    │  ← Rotes PDF-Symbol
│                 │
│  + Zahnrad ⚙️   │  ← Plus + Zahnrad für Aktion
│                 │
└─────────────────┘
```

Falls du ein echtes Icon brauchst, kann ich eines generieren. 

## Troubleshooting

| Problem | Lösung |
|---------|--------|
| Add-In wird nicht angezeigt | Pfad muss `AppData\Roaming\Microsoft\AddIns\` sein |
| Button funktioniert nicht | Sicherheitseinstellungen: **Datei** → **Optionen** → **Trust Center** → **Makro-Einstellungen** → **Alle Makros aktivieren** |
| EXE wird nicht gefunden | Excel2ZUGFeRD muss in `C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\` installiert sein |
| PDF wird nicht erstellt | Überprüfe Windows Ereignisanzeige auf Fehler |

## Schnelle Installation (Pre-built)

(Wenn wir das .xlam vorkompiliert haben, kannst du es einfach kopieren und fertig!)
