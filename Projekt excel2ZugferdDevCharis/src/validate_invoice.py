from lxml import etree
from decimal import Decimal, InvalidOperation
from datetime import datetime
import html
import os

NS = {
    "rsm": "urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100",
    "ram": "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100",
    "udt": "urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100",
}

class ValidationResult:
    def __init__(self):
        self.errors = []
        self.warnings = []

    def add_error(self, msg):
        self.errors.append(msg)

    def add_warning(self, msg):
        self.warnings.append(msg)

    def is_ok(self):
        return not self.errors

    def to_html(self, filename):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        html_content = f"""
        <html>
        <head>
            <meta charset="utf-8">
            <title>Factur-X Prüfbericht</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h1 {{ color: #333; }}
                .ok {{ color: green; font-weight: bold; }}
                .error {{ color: red; }}
                .warning {{ color: orange; }}
                .box {{ border: 1px solid #ccc; padding: 15px; margin-top: 20px; }}
            </style>
        </head>
        <body>
            <h1>Prüfbericht – Factur‑X / EN16931</h1>
            <p><strong>Datei:</strong> {html.escape(filename)}</p>
            <p><strong>Erstellt am:</strong> {timestamp}</p>
        """

        if self.errors:
            html_content += "<div class='box'><h2 class='error'>Fehler</h2><ul>"
            for e in self.errors:
                html_content += f"<li class='error'>{html.escape(e)}</li>"
            html_content += "</ul></div>"
        else:
            html_content += "<p class='ok'>Keine Fehler gefunden.</p>"

        if self.warnings:
            html_content += "<div class='box'><h2 class='warning'>Warnungen</h2><ul>"
            for w in self.warnings:
                html_content += f"<li class='warning'>{html.escape(w)}</li>"
            html_content += "</ul></div>"
        else:
            html_content += "<p class='ok'>Keine Warnungen.</p>"

        html_content += "</body></html>"

        report_name = filename + "_report.html"
        with open(report_name, "w", encoding="utf-8") as f:
            f.write(html_content)

        return report_name


def get_text(node, xpath):
    el = node.xpath(xpath, namespaces=NS)
    return el[0].text if el else None


def parse_decimal(text, default=None):
    try:
        return Decimal(text)
    except:
        return default


def validate_date_format(root, result):
    for dt in root.xpath(".//udt:DateTimeString[@format='102']", namespaces=NS):
        val = dt.text or ""
        if len(val) != 8 or not val.isdigit():
            result.add_error(f"Ungültiges Datumsformat: {val} (erwartet YYYYMMDD)")


def check_bt_presence(root, bt_name, xpath, result):
    if not root.xpath(xpath, namespaces=NS):
        result.add_error(f"{bt_name} fehlt: {xpath}")


def validate_bt_fields(root, result):
    check_bt_presence(root, "BT‑1 Rechnungsnummer",
                      "/rsm:CrossIndustryInvoice/rsm:ExchangedDocument/ram:ID", result)

    check_bt_presence(root, "BT‑2 Rechnungsdatum",
                      "//ram:IssueDateTime/udt:DateTimeString", result)

    type_code = get_text(root, "//ram:TypeCode")
    if type_code != "380":
        result.add_warning(f"BT‑3 TypeCode ist {type_code}, erwartet 380")

    check_bt_presence(root, "BT‑5 Verkäufername", "//ram:SellerTradeParty/ram:Name", result)
    check_bt_presence(root, "BT‑6 Verkäuferadresse", "//ram:SellerTradeParty/ram:PostalTradeAddress", result)
    check_bt_presence(root, "BT‑7/8 Steuernummer", "//ram:SellerTradeParty/ram:SpecifiedTaxRegistration/ram:ID", result)
    check_bt_presence(root, "BT‑9 Käufername", "//ram:BuyerTradeParty/ram:Name", result)
    check_bt_presence(root, "BT‑10 Käuferadresse", "//ram:BuyerTradeParty/ram:PostalTradeAddress", result)
    check_bt_presence(root, "BT‑11 Bruttobetrag", "//ram:GrandTotalAmount", result)
    check_bt_presence(root, "BT‑12 Währung", "//ram:InvoiceCurrencyCode", result)
    check_bt_presence(root, "BT‑13 Steuerbetrag", "//ram:TaxTotalAmount", result)
    check_bt_presence(root, "BT‑14 Steuerbasis", "//ram:TaxBasisTotalAmount", result)


def validate_totals(root, result):
    header = root.xpath("//ram:SpecifiedTradeSettlementHeaderMonetarySummation", namespaces=NS)
    if not header:
        result.add_error("HeaderSummation fehlt")
        return

    header = header[0]
    tax_basis = parse_decimal(get_text(header, "ram:TaxBasisTotalAmount"))
    tax_total = parse_decimal(get_text(header, "ram:TaxTotalAmount"))
    grand_total = parse_decimal(get_text(header, "ram:GrandTotalAmount"))

    line_totals = [
        parse_decimal(n.text, Decimal("0"))
        for n in root.xpath("//ram:IncludedSupplyChainTradeLineItem//ram:LineTotalAmount", namespaces=NS)
    ]
    sum_lines = sum(line_totals)

    if sum_lines != tax_basis:
        result.add_error(f"Positionssumme {sum_lines} != Steuerbasis {tax_basis}")

    expected_grand = (tax_basis + tax_total).quantize(Decimal("0.01"))
    if expected_grand != grand_total:
        result.add_error(f"Bruttosumme {grand_total} != Netto+Steuer {expected_grand}")


def validate_periods(root, result):
    lines = root.xpath("//ram:IncludedSupplyChainTradeLineItem", namespaces=NS)
    for line in lines:
        line_id = get_text(line, "ram:AssociatedDocumentLineDocument/ram:LineID")
        period = line.xpath(".//ram:BillingSpecifiedPeriod", namespaces=NS)
        if not period:
            result.add_warning(f"Position {line_id}: Leistungszeitraum fehlt")
        else:
            start = get_text(period[0], "ram:StartDateTime/udt:DateTimeString")
            end = get_text(period[0], "ram:EndDateTime/udt:DateTimeString")
            if not start or not end:
                result.add_warning(f"Position {line_id}: Start/Ende unvollständig")


def validate_profile(root, result):
    profile = get_text(root, "//ram:GuidelineSpecifiedDocumentContextParameter/ram:ID")
    if not profile:
        result.add_error("Profil‑ID fehlt")
    elif "en16931" not in profile.lower():
        result.add_warning(f"Profil wirkt nicht EN16931‑konform: {profile}")


def validate_invoice(path):
    result = ValidationResult()

    try:
        tree = etree.parse(path)
        root = tree.getroot()
    except Exception as e:
        result.add_error(f"XML‑Fehler: {e}")
        return result

    validate_date_format(root, result)
    validate_bt_fields(root, result)
    validate_profile(root, result)
    validate_totals(root, result)
    validate_periods(root, result)

    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python validate_invoice.py factur-x.xml")
        sys.exit(1)

    xml_path = sys.argv[1]
    res = validate_invoice(xml_path)
    print(res)

    report = res.to_html(xml_path)
    print(f"\nHTML‑Report erstellt: {report}")
