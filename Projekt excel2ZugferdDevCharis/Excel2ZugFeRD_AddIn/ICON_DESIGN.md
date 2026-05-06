# 🎨 Icon Design für "ZgFeRD pdf erstellen" Button

## Icon Vorschlag #1: Minimalistisch (Empfohlen)

```
128x128 Pixel - Transparenter Hintergrund

     ┏━━━━━━━━━━━┓
     ┃ ╔═══════╗ ┃
     ┃ ║ P D F ║ ┃ ← Rotes PDF-Dokument
     ┃ ║   +   ║ ┃   (RGB: 192, 0, 0)
     ┃ ║  ⚙️   ║ ┃
     ┃ ╚═══════╝ ┃
     ┗━━━━━━━━━━━┛

Größeverhältnis:
- PDF Dokument: 80x90px, Rot (#C00000)
- Zahnrad: 40x40px, Blau (#1F4788), Bottom-Right Corner
- Positionierung: Zentralisiert
```

---

## Icon Vorschlag #2: Mit Text

```
┌────────────────────┐
│   PDF  PDF  PDF    │ ← 3x "PDF" übereinander = Dokument-Stack
│                    │
│   ⚙️ ZUGFeRD       │ ← Zahnrad + Text
│                    │
└────────────────────┘
```

---

## Icon Vorschlag #3: Rechnung + PDF

```
Kombination aus:
- 📄 Dokument-Symbol (oben)
- 🔒 Schloss (unten rechts) = Sicherheit (ZUGFeRD ist signiert)
- € oder 💰 = Rechnung/Zahlungsbezug
```

---

## Farben (RGB)

| Element | RGB | HEX |
|---------|-----|-----|
| PDF-Rot | 192, 0, 0 | #C00000 |
| Blau (Zahnrad) | 31, 71, 136 | #1F4788 |
| Grün (Checkmark) | 0, 176, 80 | #00B050 |
| Grau (Hintergrund) | 242, 242, 242 | #F2F2F2 |
| Weiß (Text) | 255, 255, 255 | #FFFFFF |

---

## Button Label

**Text:** `ZgFeRD pdf erstellen`

**Schriftart:** Segoe UI (Windows Standard)
**Größe:** 11pt
**Fett:** Ja
**Farbe:** Schwarz oder Dunkelblau

---

## Größen

Für verschiedene Excel Button Größen:

| Größe | Pixel |
|-------|-------|
| Klein | 16x16 |
| Normal | 32x32 |
| Groß | 64x64 |
| Sehr Groß | 128x128 |

**Tipp:** Speichere in **32x32 oder 128x128** für beste Qualität

---

## Icon als PNG Generator

Falls du ein echtes Icon brauchst, kann ich dieses Python-Skript verwenden:

```python
from PIL import Image, ImageDraw, ImageFont

# 128x128 PDF Icon mit Zahnrad
def create_icon():
    img = Image.new('RGBA', (128, 128), (242, 242, 242, 255))
    draw = ImageDraw.Draw(img)
    
    # PDF Dokument (Rot)
    pdf_color = (192, 0, 0, 255)
    draw.rectangle([20, 15, 85, 100], fill=pdf_color, outline=(0, 0, 0, 255), width=2)
    draw.text((35, 40), "PDF", fill=(255, 255, 255), font=None)
    
    # Zahnrad (Blau) - Bottom Right
    gear_color = (31, 71, 136, 255)
    draw.ellipse([75, 80, 110, 115], fill=gear_color, outline=(0, 0, 0, 255), width=1)
    draw.text((82, 92), "⚙", fill=(255, 255, 255), font=None)
    
    img.save('zugferd_icon.png')

create_icon()
```

---

## Recommendation

✅ **Use Icon #1 (Minimalistisch)** - 
- Sauberes, modernes Design
- Klar erkennbar
- Professionell
- Beim Button wird es sowieso klein dargestellt

**Farben:**
- 🔴 PDF: Rot (#C00000) - kommt sofort ins Auge
- ⚙️ Zahnrad: Dunkelblau (#1F4788) - Verarbeitung/Aktion

---

## Speichern

Falls du das Icon speichern möchtest:
1. Kopiere den visuellen Design oben
2. Oder nutze einen Online Icon Generator
3. Speichere als **PNG, 128x128px**, transparenter Hintergrund
4. Benenne die Datei: `icon_zugferd.png`

Das Icon wird vom Excel Button automatisch skaliert!
