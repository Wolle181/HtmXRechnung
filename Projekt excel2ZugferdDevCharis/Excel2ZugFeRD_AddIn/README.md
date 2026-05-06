# Excel2ZugFeRD Excel Add-In

Dieses Add-In ermöglicht es, direkt aus Excel heraus ZUGFeRD-konforme PDF-Rechnungen zu erstellen.

## 📦 Inhalt

- **Schnell_Add-In.bas** - VBA Code (copy-paste ready)
- **Module_ZugFeRD.bas** - Erweiterte VBA mit Error Handling
- **customUI.xml** - Ribbon Definition (optional)
- **SCHNELLSTART.md** - 5-Minuten Installationsanleitung ⭐ START HERE
- **ANLEITUNG.md** - Detaillierte Anleitung mit Troubleshooting
- **ICON_DESIGN.md** - Icon Vorschläge und Design

## ⚡ TL;DR (30 Sekunden)

1. `SCHNELLSTART.md` öffnen & folgen (5 Minuten)
2. VBA-Code aus `Schnell_Add-In.bas` kopieren
3. Button hinzufügen
4. Speichern als `.xlam` in AddIns Folder
5. Fertig! ✅

## 🎯 Was macht das Add-In?

- ✅ Button "ZgFeRD pdf erstellen" im Excel-Menü
- ✅ Ruft Excel2ZugFeRD Anwendung auf
- ✅ Erstellt ZUGFeRD-konforme PDF-Rechnung
- ✅ Error Handling & User Feedback
- ✅ Automatische Blattauswahl

## ✨ Features

**VBA Features:**
- Arbeitsmappe automatisch speichern
- Fehlerbehandlung & Validierung
- Sicherheitschecks
- Benutzer-Feedback (Meldungen)
- Automatische Blattindexberechnung
- Event Viewer Integration

**UI Features:**
- Button mit Text "ZgFeRD pdf erstellen"
- Icon (optional, Vorschlag in ICON_DESIGN.md)
- Im Home-Tab platzierbar
- Oder in eigenem Tab

## 📋 System-Anforderungen

- ✅ Excel 2016 oder neuer
- ✅ VBA aktiviert
- ✅ Excel2ZugFeRD installiert (`C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\`)
- ✅ Windows (für Shell.Run())

## 🔒 Sicherheit

- ✅ Makro-Sicherheit kann aktiviert werden
- ✅ Nur lokale Ausführung (keine Internet-Verbindung)
- ✅ Quellcode transparent (VBA ist readable)
- ✅ Keine Daten-Sammlung

## 📖 Installation

**WICHTIG:** Siehe `SCHNELLSTART.md` für schritt-für-schritt Anleitung!

Kurz:
1. VBA-Code in neues Modul einfügen
2. Button auf Sheet hinzufügen
3. Als `.xlam` speichern in:
   ```
   C:\Users\[USERNAME]\AppData\Roaming\Microsoft\AddIns\
   ```
4. In Excel aktivieren: Datei → Optionen → Add-Ins → Durchsuchen

## 🚀 Nutzung

1. Excel öffnen
2. Button "ZgFeRD pdf erstellen" klicken
3. Excel2ZugFeRD öffnet sich
4. PDF wird erstellt ✅

## 🐛 Troubleshooting

| Problem | Lösung |
|---------|--------|
| Add-In nicht sichtbar | Siehe SCHNELLSTART.md Schritt 6 |
| Button funktioniert nicht | Makro-Sicherheit aktivieren (Schritt oben) |
| "Datei nicht gefunden" | Excel2ZugFeRD installiert? Siehe ANLEITUNG.md |
| PDF wird nicht erstellt | Event Viewer prüfen, Excel-Struktur validieren |

**Weitere Hilfe:** `ANLEITUNG.md` → Troubleshooting Section

## 📝 Customization

### Button Text ändern
In `Schnell_Add-In.bas` Zeile ändern:
```vba
' Ändere zu deinem Text:
MsgBox msg, vbInformation, "Dein Text hier"
```

### Blatt-Index ändern
```vba
Const BLATT_NR As String = "0"  ' 0=1.Blatt, 1=2.Blatt, etc.
```

### EXE-Pfad ändern
```vba
Const EXCEL2ZUGFERD_EXE As String = "C:\Dein\Pfad\excel2zugferd.exe"
```

## 🎨 Icon hinzufügen

1. PNG Icon erstellen (128x128, transparent)
2. Icon in Add-In einbetten (komplexer - siehe `customUI.xml`)
3. Oder: Button ohne Icon nutzen (einfacher)

Siehe `ICON_DESIGN.md` für Vorschläge

## 📚 Weiterführende Ressourcen

- **VBA Dokumentation:** https://docs.microsoft.com/office/vba/
- **Ribbon XML:** https://docs.microsoft.com/office/client-developer/
- **Excel2ZugFeRD:** https://github.com/Lkammer/excel2zugferd

## 📞 Support

Probleme?
1. **SCHNELLSTART.md** lesen ← Start hier!
2. **ANLEITUNG.md** → Troubleshooting
3. **Event Viewer** überprüfen
4. GitHub Issues: https://github.com/Lkammer/excel2zugferd/issues

## 📄 Lizenz

Same as Excel2ZugFeRD (siehe hauptes Repositories)

---

**Version:** 1.0  
**Letztes Update:** 2026  
**Status:** Ready to use ✅
