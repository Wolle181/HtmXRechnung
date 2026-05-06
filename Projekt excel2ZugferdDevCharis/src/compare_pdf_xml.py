import pdfplumber
from lxml import etree
from src.validate_invoice import get_text, NS
import locale


def extract_pdf_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def _to_german_number(value: str) -> str:
    """Konvertiert '1234.56' → '1.234,56' für Vergleich mit PDF-Text."""
    try:
        f = float(value)
        # Deutsch formatiert mit Tausender-Punkt und Komma als Dezimalzeichen
        return f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return value


def compare_pdf_xml(pdf_path, xml_path):
    result = {
        "errors": [],
        "warnings": [],
        "matches": []
    }

    # PDF lesen
    pdf_text = extract_pdf_text(pdf_path)

    # XML lesen
    tree = etree.parse(xml_path)
    root = tree.getroot()

    # Rechnungsnummer: BT-1 liegt unter ExchangedDocument/ram:ID
    rg_nr = get_text(root, "/rsm:CrossIndustryInvoice/rsm:ExchangedDocument/ram:ID")

    # Datum: YYYYMMDD → TT.MM.JJJJ
    datum_raw = get_text(root, "//ram:IssueDateTime/udt:DateTimeString")
    datum_pdf = None
    if datum_raw and len(datum_raw) == 8:
        datum_pdf = f"{datum_raw[6:8]}.{datum_raw[4:6]}.{datum_raw[0:4]}"

    brutto_raw = get_text(root, "//ram:GrandTotalAmount")
    steuer_raw = get_text(root, "//ram:TaxTotalAmount")
    netto_raw  = get_text(root, "//ram:TaxBasisTotalAmount")

    checks = {
        "Rechnungsnummer": rg_nr,
        "Rechnungsdatum":  datum_pdf,
        "Bruttobetrag":    _to_german_number(brutto_raw),
        "Steuerbetrag":    _to_german_number(steuer_raw),
        "Nettobetrag":     _to_german_number(netto_raw),
    }

    for label, value in checks.items():
        if value and value in pdf_text:
            result["matches"].append(f"{label}: {value}")
        else:
            result["errors"].append(f"{label}: '{value}' nicht im PDF gefunden")

    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Verwendung: compare_pdf_xml.py <pdf-datei> <xml-datei>")
        sys.exit(1)
    pdf_path, xml_path = sys.argv[1], sys.argv[2]
    result = compare_pdf_xml(pdf_path, xml_path)
    if result["matches"]:
        print("=== Übereinstimmungen ===")
        for m in result["matches"]:
            print(f"  OK  {m}")
    if result["warnings"]:
        print("=== Warnungen ===")
        for w in result["warnings"]:
            print(f"  !!  {w}")
    if result["errors"]:
        print("=== Fehler ===")
        for e in result["errors"]:
            print(f"  XX  {e}")

