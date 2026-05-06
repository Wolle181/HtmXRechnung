"""
Löst das eingebettete ZUGFeRD/Factur-X XML aus einer PDF-Rechnung heraus
und speichert es neben der PDF-Datei.

Aufruf:
    uv run python -m src.extract_xml_from_pdf "Pfad/zur/Rechnung.pdf"

Ausgabe:
    <gleicher Ordner>/factur-x.xml   (oder der Name des eingebetteten Anhangs)
"""

import sys
from pathlib import Path
from pypdf import PdfReader


def extract_xml(pdf_path: str) -> str:
    """Extrahiert die erste XML-Anlage aus pdf_path und gibt den Ausgabepfad zurück."""
    pdf_path = Path(pdf_path)
    reader = PdfReader(str(pdf_path))

    # Eingebettete Dateien liegen unter /EmbeddedFiles im PDF-Katalog
    embedded = reader.trailer["/Root"].get("/Names", {}).get("/EmbeddedFiles", {})
    names_array = embedded.get("/Names", [])

    # names_array enthält abwechselnd Name und Objekt-Referenz
    pairs = list(zip(names_array[0::2], names_array[1::2]))

    xml_pairs = [(name, ref) for name, ref in pairs
                 if str(name).lower().endswith(".xml")]

    if not xml_pairs:
        raise ValueError("Keine XML-Anlage in der PDF gefunden.")

    # Ersten XML-Anhang verwenden
    file_name, file_ref = xml_pairs[0]
    ef = file_ref.get_object()
    file_spec = ef.get("/EF", ef)
    stream = file_spec.get("/F") or file_spec.get("/UF")
    if stream is None:
        raise ValueError("XML-Stream konnte nicht gelesen werden.")

    xml_data = stream.get_object().get_data()

    out_path = pdf_path.parent / Path(str(file_name)).name
    out_path.write_bytes(xml_data)
    return str(out_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Verwendung: uv run python -m src.extract_xml_from_pdf <PDF-Datei>")
        sys.exit(1)

    try:
        result = extract_xml(sys.argv[1])
        print(f"XML gespeichert: {result}")
    except Exception as e:
        print(f"Fehler: {e}")
        sys.exit(1)
