# 🎯 Excel Add-In - SCHNELLSTART (5 Minuten)

## ⚡ Die schnellste Methode

Falls du KEINE Custom Ribbon brauchst (nur einen Button), verwende diese Methode:

### Schritt-für-Schritt:

#### 1️⃣ Excel öffnen - Neue Datei
```
Datei → Neu → Leere Arbeitsmappe
```

#### 2️⃣ Developer Mode aktivieren (falls noch nicht)
```
Datei → Optionen → Menüleiste anpassen
→ Haken bei "Entwickler"
```

#### 3️⃣ VBA-Code einfügen
```
Entwickler → Code anzeigen (oder Alt+F11)
→ Rechtsklick auf "VBAProject" 
→ "Modul einfügen"
→ Code aus "Schnell_Add-In.bas" einführgen
→ Speichern (Ctrl+S)
```

#### 4️⃣ Button hinzufügen
```
Entwickler → Einfügungssteuerelemente 
→ "Schaltfläche" (FormControl)
→ Rechteck auf dem Sheet zeichnen
→ Dialog: Macro wählen → "CreateZugFeRDPDF"
→ OK
→ Button-Text ändern: "ZgFeRD pdf erstellen"
```

#### 5️⃣ Speichern als Add-In
```
Datei → Speichern unter
→ Dateityp: "Excel-Add-In (*.xlam)"
→ Name: "Excel2ZugFeRD_AddIn"
→ Speicherort: 
   C:\Users\[DEIN_NAME]\AppData\Roaming\Microsoft\AddIns\
→ Speichern → Fertig!
```

#### 6️⃣ Aktivieren in Excel
```
Datei → Optionen → Add-Ins
→ Verwalten: "Excel-Add-Ins"
→ "Durchsuchen..."
→ "Excel2ZugFeRD_AddIn.xlam" auswählen
→ OK
```

#### 7️⃣ Testen
```
Neue Arbeitsmappe öffnen
→ Button sollte sichtbar sein
→ Klick auf Button → ZUGFeRD PDF wird erstellt!
```

---

## 🎨 Icon Vorschlag

```
┌──────────────────────────┐
│                          │
│        ┌─────┐           │
│        │ PDF │           │ ← Rotes PDF-Dokument
│        │  +  │           │   mit Zahnrad überlagert
│        │  ⚙️  │           │
│        └─────┘           │
│                          │
│  "ZgFeRD pdf erstellen"  │ ← Button Text
│                          │
└──────────────────────────┘

Farben:
- Hintergrund: Rot (#C00000) für PDF
- Zahnrad: Dunkelblau (#1F4788)
- Text: Weiß
- Icons: Modern, minimal
```

**Icon Generator Option:**
Falls du ein echtes Icon brauchst, kann ich eins mit Python generieren:
```python
from PIL import Image, ImageDraw, ImageFont

# 128x128px Icon mit:
# - Rotes PDF-Symbol
# - Zahnrad-Overlay
# - Text darunter
```

---

## 🚨 Troubleshooting

### Problem: Add-In wird in Excel nicht angezeigt

**Lösung:**
1. **Pfad kontrollieren** - Muss exakt sein:
   ```
   C:\Users\[DEIN_BENUTZERNAME]\AppData\Roaming\Microsoft\AddIns\
   ```
   
2. **AppData ist versteckt** - Muss aktiviert sein:
   - Dateie-Explorer → Ansicht → Optionen
   - Haken bei "Versteckte Dateien anzeigen"

3. **Excel neu starten** nach dem Speichern

---

### Problem: Button funktioniert nicht / Fehler

**Lösung:**
1. **Sicherheitseinstellungen anpassen:**
   ```
   Datei → Optionen → Trust Center → Trust Center-Einstellungen
   → Makro-Einstellungen
   → Haken bei "Alle Makros aktivieren" (mit Benachrichtigungen)
   ```

2. **Excel neu starten**

3. **Fehlerlog prüfen:**
   - Windows Taskeiste → Rechtsklick auf Start → Ereignisanzeige
   - Windows-Protokolle → Anwendung
   - Suche nach "Excel2ZugFeRD" Fehlern

---

### Problem: "Die Anwendung Excel2ZugFeRD wurde nicht gefunden"

**Lösung:**
- Excel2ZugFeRD muss installiert sein:
  ```
  https://github.com/Lkammer/excel2zugferd/releases
  ```
- Standard Installationspfad:
  ```
  C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\excel2zugferd.exe
  ```

---

### Problem: PDF wird nicht erstellt / "Exit Code: -1"

**Lösung:**
1. **Arbeitsmappe speichern** (im Format .xlsx oder .xlsm)
2. **Excel-Struktur überprüfen:**
   - Erste Zeile: Überschriften
   - Spalten korrekt? (Pos, Datum, Beschreibung, etc.)
   - Keine leeren Zeilen in den Daten

3. **Windows Event Viewer überprüfen:**
   ```
   Ereignisanzeige → Windows-Protokolle → Anwendung
   → "Excel2ZugFeRD" Fehler-Einträge
   ```

---

## 📝 Weitere Optionen

### Mit Ribbon im Home-Tab (Professionell)
Wenn du willst, dass der Button im **Home-Tab** (oben links) sichtbar ist:
- Braucht mehr Setup (XML-Ribbon)
- Siehe: ANLEITUNG.md für erweiterte Version

### Mit Ribbon im eigenen Tab
- Noch professioneller
- Braucht Custom Ribbon Definition
- Komplexer zum Setup

---

## ✅ Alles Fertig!

Dein Excel Add-In ist jetzt aktiviert und bereit für die Nutzung.

**Nächste Schritte:**
1. Excel öffnen
2. Neue Arbeitsmappe oder bestehende öffnen
3. Button "ZgFeRD pdf erstellen" klicken
4. ZUGFeRD PDF wird automatisch erstellt! 🎉

---

## 📞 Support

Falls Fehler auftreten:
- Check Event Viewer (oben)
- Überprüfe Excel-Datei Struktur
- Siehe GitHub Issues: https://github.com/Lkammer/excel2zugferd/issues
