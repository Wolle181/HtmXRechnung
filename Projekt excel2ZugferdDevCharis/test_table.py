from fpdf import FPDF

# Beispiel: 8 Spalten, Spalte H gefüllt
data = [
    ["Pos.", "Datum", "Beschreibung", "Menge", "Einheit", "Einzelpreis", "Gesamt", "H"],
    ["1", "2024-01-01", "Test", "1", "Stk", "10", "10", "Bemerkung H"],
    ["2", "2024-01-02", "Test2", "2", "Stk", "20", "40", ""]
]

class SimplePdf(FPDF):
    pass

pdf = SimplePdf()
pdf.add_page()
pdf.set_font("Arial", size=10)

# Spaltenbreiten für 8 Spalten
col_widths = [15, 25, 50, 15, 20, 20, 20, 30]

# Header
for i, header in enumerate(data[0]):
    pdf.cell(col_widths[i], 10, header, border=1)
pdf.ln()
# Daten
for row in data[1:]:
    for i, cell in enumerate(row):
        pdf.cell(col_widths[i], 10, str(cell), border=1)
    pdf.ln()

pdf.output("test.pdf")
print("PDF test.pdf wurde erzeugt.")
from src.handle_pdf import Pdf

data = [
    ["Pos.", "Datum", "Beschreibung", "Menge", "Einheit", "Einzelpreis", "Gesamt", "H"],
    ["1", "2024-01-01", "Test", "1", "Stk", "10", "10", "Bemerkung H"],
    ["2", "2024-01-02", "Test2", "2", "Stk", "20", "40", ""]
]

pdf = Pdf()
pdf.add_page()
pdf.print_positions(data)
pdf.output("test.pdf")
print("PDF test.pdf wurde erzeugt.")
