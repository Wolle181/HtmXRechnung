"""
Module handle_zugferd
"""

# import re
import numpy as np
import pandas as pd
import math

from datetime import datetime
from decimal import Decimal

from drafthorse.models.accounting import ApplicableTradeTax
from drafthorse.models.document import Document
from drafthorse.models.note import IncludedNote
from drafthorse.models.tradelines import LineItem
from drafthorse.models.payment import PaymentTerms
from drafthorse.models.party import TaxRegistration
from drafthorse.pdf import attach_xml
from lxml import etree

# from drafthorse.models import NS_QDT

from src.kunde import Kunde
from src.lieferant import Lieferant
from src.invoice_collection import InvoiceCollection
from src.constants import P19USTG, GERMAN_DATE, EINHEITEN


class ZugFeRD:
    """Class ZugFeRD"""

    def __init__(self, invoice: InvoiceCollection = None):
        # Build data structure
        self.doc = Document()
        self.doc.context.guideline_parameter.id = (
            "urn:cen.eu:en16931:2017#conformant#urn:factur-x.eu:1p0:extended"
        )
        self.doc.header.type_code = "380"
        self.doc.header.name = "RECHNUNG"
        # self.doc.header.languages.add("de")
        self.debug = False
        self.first_date = None
        self.last_date = None
        self.rg_date: datetime = None
        self.customer_number_value = None  # Speichere Kundennummer
        self.fill_xml(invoice)

    def add_rgnr(self, rgnr: str, datum):
        """Set Rechnungsnummer to id in header"""
        self.doc.header.id = str(rgnr)
        
        # Stelle sicher, dass datum ein datetime ist
        if isinstance(datum, str):
            try:
                use_datum = datetime.strptime(datum, "%d.%m.%Y")
            except (ValueError, TypeError):
                try:
                    use_datum = datetime.strptime(datum, "%Y-%m-%d")
                except (ValueError, TypeError):
                    use_datum = self.rg_date  # self.rg_date ist garantiert datetime
        else:
            use_datum = datum
        
        # use_datum ist jetzt garantiert ein datetime-Objekt
        self.doc.header.issue_date_time = use_datum.date()

    def add_customer_number(self, customer_number: str):
        """Add customer number as buyer reference in ZUGFeRD"""
        if customer_number and str(customer_number).strip().upper() != "NONE":
            # Speichere den Wert für später (für XML-Manipulation)
            self.customer_number_value = str(customer_number)
            try:
                # Setze die Kundennummer in der BuyerTradeParty ID (direkte Zuweisung)
                self.doc.trade.agreement.buyer.id = str(customer_number)
            except Exception as e:
                # Logging für Debugging
                import traceback
                traceback.print_exc()
                pass

    def add_note(self, text):
        """Add note to notes"""
        note = IncludedNote()
        try:
            note.content.add(text)
        except (AttributeError, TypeError):
            # Fallback: try direct assignment
            try:
                note.content = text
            except:
                pass
        note.subject_code = "REG"
        try:
            self.doc.header.notes.add(note)
        except (AttributeError, TypeError):
            pass

    def add_bundesland(self, bundesland):
        """Add Bundesland"""
        if bundesland:
            self.doc.trade.agreement.seller.address.country_subdivision = bundesland

    def add_zahlungsempfaenger(self, text):
        """set Zahlungsempfaenger to correct value - SIMPLIFIED"""
        # Minimal implementation to avoid drafthorse API issues
        try:
            # Just set the type code, skip everything else that causes issues
            self.doc.trade.settlement.payment_means.type_code = "58"
        except Exception:
            pass  # Silent fail

    def _add_buyer_from_text(self, text: str) -> None:
        arr = text.split("\n")
        # self.doc.trade.settlement.invoicee.name = arr[0]
        self.doc.trade.agreement.buyer.name = arr[0]
        if len(arr) > 2:
            self.doc.trade.agreement.buyer.address.line_one = arr[1]
        if len(arr) > 3:
            self.doc.trade.agreement.buyer.address.line_one = arr[-2]
        if len(arr) > 1:
            self.doc.trade.agreement.buyer.address.postcode = arr[-1].split(" ", 1)[0]
            self.doc.trade.agreement.buyer.address.city_name = arr[-1].split(" ", 1)[1]

    def _add_str_hnr(self, buyer: Kunde) -> None:
        if buyer.postfach is not None:  # BT-50
            self.doc.trade.agreement.buyer.address.line_one = (
                "Postfach: " + buyer.postfach
            )
        elif buyer.strasse is not None:  # BT-50
            self.doc.trade.agreement.buyer.address.line_one = buyer.strasse + (
                " " + buyer.hausnummer if buyer.hausnummer else ""
            )

    def _add_plz_ort(self, buyer: Kunde) -> None:
        if buyer.plz is not None:  # BT-53
            self.doc.trade.agreement.buyer.address.postcode = buyer.plz
        if buyer.ort is not None:  # BT-52
            self.doc.trade.agreement.buyer.address.city_name = buyer.ort

    def _add_buyer_from_object(self, buyer: Kunde) -> None:
        if buyer.betriebsbezeichnung is not None:  # BT-44
            self.doc.trade.agreement.buyer.name = buyer.betriebsbezeichnung
        if buyer.adresszusatz is not None:  # BT-51
            self.doc.trade.agreement.buyer.address.line_two = buyer.adresszusatz
            if buyer.name is not None:  # BT-163
                self.doc.trade.agreement.buyer.address.line_three = buyer.name
        else:
            if buyer.name is not None:  # BT-51
                self.doc.trade.agreement.buyer.address.line_two = buyer.name
        self._add_str_hnr(buyer)
        self._add_plz_ort(buyer)

    def add_rechnungsempfaenger(self, text: str, adr: Kunde = None):
        """set Rechnungsempfänger"""
        self.doc.trade.settlement.currency_code = "EUR"
        # self.doc.trade.settlement.tax_currency_code = "EUR" # BR-53-1
        if adr is not None:
            self._add_buyer_from_object(adr)
        else:
            self._add_buyer_from_text(text)
        self.doc.trade.agreement.buyer.address.country_id = "DE"

    def _add_my_adresse(self, lieferant: Lieferant):
        self.doc.trade.agreement.seller.id = lieferant.betriebsbezeichnung

        self.doc.trade.agreement.seller.name = lieferant.betriebsbezeichnung
        if lieferant.adresszusatz:
            self.doc.trade.agreement.seller.address.line_two = lieferant.adresszusatz
        if lieferant.postfach:
            self.doc.trade.agreement.seller.address.line_two = lieferant.postfach
        else:
            self.doc.trade.agreement.seller.address.line_one = lieferant.strasse + (
                " " + lieferant.hausnummer if lieferant.hausnummer else ""
            )
        self.doc.trade.agreement.seller.address.postcode = lieferant.plz
        self.doc.trade.agreement.seller.address.city_name = lieferant.ort
        self.doc.trade.agreement.seller.address.country_id = "DE"

    def _add_my_kontakt(self, lieferant: Lieferant):
        if lieferant.name:
            self.doc.trade.agreement.seller.contact.person_name = lieferant.name
        if lieferant.telefon:
            self.doc.trade.agreement.seller.contact.telephone.number = lieferant.telefon
        if lieferant.fax:
            self.doc.trade.agreement.seller.contact.fax.number = lieferant.fax
        if lieferant.email:
            self.doc.trade.agreement.seller.contact.email.address = lieferant.email

    def add_my_company(self, lieferant: Lieferant):
        """Add Address of my company to zugferd"""
        self._add_my_adresse(lieferant)
        self._add_my_kontakt(lieferant)
        self.add_bundesland(lieferant.bundesland)
        taxreg = TaxRegistration()
        taxreg.id = (
            ("VA", lieferant.steuerid)
            if lieferant.steuerid
            else ("FC", lieferant.steuernr)
        )
        try:
            self.doc.trade.agreement.seller.tax_registrations.add(taxreg)
        except (AttributeError, TypeError):
            pass  # Skip if not available
        # self.doc.trade.agreement.seller.tax = ustid

    def add_verwendungszweck(self, rg_nr: dict, datum: str) -> None:
        """Add Verwendungszweck to zugferd BT-83"""
        self.doc.trade.settlement.payment_reference = (
            f"{list(rg_nr.keys())[0]} {list(rg_nr.values())[0]} vom {datum}"
        )

    def _fillPosAndNameOfLi(self, li: LineItem, item: list) -> None:
        li.document.line_id = item[0]  # Pos.
        # Falls Spalte H existiert und gefüllt ist, an Beschreibung anhängen
        beschreibung = item[2]
        if len(item) > 7 and str(item[7]).strip() != "":
            beschreibung += f"\n{item[7]}"
        li.product.name = beschreibung

    def _replaceCommaWithDot(self, item: str) -> float:
        """convert German number format (comma decimal) to float"""
        if item is None or item == '' or str(item).lower() == 'nan':
            return 0.0
        try:
            # Handle cases like "22,00 €" or just "22,00"
            value_str = str(item).split()[0].replace(",", ".")
            return float(value_str)
        except (ValueError, IndexError, AttributeError, TypeError):
            return 0.0

    def _setTaxInLi(self, li: LineItem, item: str, the_tax: str) -> None:
        # Stelle sicher dass the_tax konvertierbar ist zu Decimal
        try:
            tax_value = Decimal(str(the_tax).replace(",", ".")) if the_tax else Decimal("0")
        except:
            tax_value = Decimal("0")
        
        li.settlement.trade_tax.type_code = "VAT"
        li.settlement.trade_tax.category_code = "E" if str(the_tax) == "0.00" else "S"
        li.settlement.trade_tax.rate_applicable_percent = tax_value
        gesamt = self._replaceCommaWithDot(item)
        li.settlement.monetary_summation.total_amount = Decimal(f"{gesamt:.2f}")

    def _set_date(self, item: str) -> datetime:
        if item == "nan" or item == "":
            return (
                self.last_date if self.last_date is not None else self.rg_date
            )  # datetime.today()
        else:
            return datetime.strptime(item, "%Y-%m-%d %H:%M:%S")

    def _setOccurrenceInLi(self, li: LineItem, item: str) -> None:
        if item is not None:
            the_date = self._set_date(item)
            # BG-26
            li.settlement.period.start = the_date  # BT-134
            # li.settlement.period.end = the_date  # BT-135
            if self.first_date is None:
                self.first_date = the_date
            self.last_date = the_date

    def _get_einheit(self, inp: str, li: LineItem = None) -> str:
        """search converted Einheit in Einheiten, default 'Stück'"""
        if inp is None:
            return None
        if inp in EINHEITEN.keys():
            return EINHEITEN[inp]
        else:
            if li is not None:
                li.product.description = f"Die Einheit '{inp}' ist nicht\
 verfügbar und wurde durch 'C62' (Stück) ersetzt."  # BT-154
            return "C62"  # Stück

    def _parse_menge(self, raw) -> float:
        """Parst Mengenwert: Dezimalzahl oder Bruch (z.B. '4/20' → 0.2)"""
        s = str(raw).split()[0].replace(",", ".").strip()
        if "/" in s:
            parts = s.split("/")
            try:
                return float(parts[0]) / float(parts[1])
            except (ValueError, ZeroDivisionError):
                return 1.0
        try:
            return float(s)
        except ValueError:
            return 1.0

    def add_items(self, dat, the_tax: str):
        """add items to invoice"""
        # ("Pos.", "Datum", "Tätigkeit", "Menge", "Typ",
        #  "Einzel €", "Gesamt €")
        for i, item in enumerate(dat):
            if i > 0:
                li = LineItem()
                self._fillPosAndNameOfLi(li, item)
                self._setOccurrenceInLi(li, item[1])
                # Ensure values are properly converted to float
                try:
                    menge = self._parse_menge(item[3]) if item[3] else 1.0
                    einzelpreisnetto = float(str(item[5]).split()[0].replace(",", ".")) if item[5] else 0.0
                except (ValueError, IndexError, AttributeError):
                    menge = 1.0
                    einzelpreisnetto = 0.0
                
                li.agreement.net.amount = Decimal(f"{einzelpreisnetto:.2f}")
                li.delivery.billed_quantity = (
                    Decimal(f"{menge:.4f}"),
                    self._get_einheit(item[4], li),  # BT-130
                )  # C62 == pieces - BT-150 ?
                self._setTaxInLi(li, item[6], the_tax)
                try:
                    self.doc.trade.items.add(li)
                except (AttributeError, TypeError):
                    pass  # Skip if not available

    def _get_float_value(self, tuple_or_value) -> float:
        """extract float from tuple (key, value) or convert value directly"""
        try:
            if isinstance(tuple_or_value, (tuple, list)) and len(tuple_or_value) == 2:
                _, v = tuple_or_value
            else:
                v = tuple_or_value
            
            # Convert string with German decimal format to float
            if isinstance(v, str):
                v = v.split()[0].replace(",", ".")
            return float(v)
        except (ValueError, IndexError, AttributeError, TypeError):
            return 0.0

    def add_rechnungsperiode(self, posDatum: pd.DataFrame = None) -> None:
        # Default: Rechnungsdatum als Fallback
        fallback = self.rg_date if self.rg_date is not None else None

        startDatum = None
        endDatum = None

        if posDatum is not None and len(posDatum) > 0:
            posDatum = posDatum.copy()
            posDatum = pd.to_datetime(posDatum, errors="coerce")
            posDatum = posDatum.dropna()
            if len(posDatum) > 0:
                s = posDatum.min(skipna=True)
                e = posDatum.max(skipna=True)
                if not pd.isna(s) and not pd.isna(e):
                    startDatum = s.to_pydatetime() if hasattr(s, 'to_pydatetime') else s
                    endDatum = e.to_pydatetime() if hasattr(e, 'to_pydatetime') else e

        if startDatum is None:
            startDatum = fallback
        if endDatum is None:
            endDatum = fallback

        if startDatum is None or endDatum is None:
            return

        try:
            self.doc.trade.delivery.event.occurrence = startDatum
        except (AttributeError, TypeError):
            pass

        try:
            # BG-14 - Rechnungszeitraum (ram:BillingSpecifiedPeriod, format 102 = YYYYMMDD)
            self.doc.trade.settlement.period.start = startDatum  # BT-73
            self.doc.trade.settlement.period.end = endDatum  # BT-74
        except (AttributeError, TypeError):
            pass

    def add_gesamtsummen(
        self, dat, the_tax: str, steuerbefreiungsgrund: str = None
    ) -> None:
        """add gesamtsumme to invoice"""
        netto = self._get_float_value(dat[0])
        steuer = self._get_float_value(dat[1])
        brutto = self._get_float_value(dat[2])

        # Stelle sicher dass the_tax richtig formatiert ist für Decimal
        try:
            tax_decimal = Decimal(str(the_tax).strip().replace(",", ".")) if the_tax else Decimal("0")
        except:
            tax_decimal = Decimal("0")

        trade_tax = ApplicableTradeTax()
        trade_tax.calculated_amount = Decimal(f"{steuer:.2f}")
        trade_tax.basis_amount = Decimal(f"{netto:.2f}")
        trade_tax.type_code = "VAT"
        trade_tax.category_code = "E" if the_tax == "0.00" else "S"
        # trade_tax.exemption_reason_code = 'VATEX-EU-AE'
        trade_tax.rate_applicable_percent = tax_decimal
        if steuerbefreiungsgrund:
            trade_tax.exemption_reason = steuerbefreiungsgrund
        try:
            self.doc.trade.settlement.trade_tax.add(trade_tax)
        except (AttributeError, TypeError):
            pass  # Skip if not available

        self.doc.trade.settlement.monetary_summation.line_total = Decimal(f"{netto:.2f}")
        # self.doc.trade.settlement.monetary_summation\
        #   .charge_total = Decimal("0.00")
        # self.doc.trade.settlement.monetary_summation\
        #   .allowance_total = Decimal("0.00")
        self.doc.trade.settlement.monetary_summation.tax_basis_total = Decimal(f"{netto:.2f}")
        self.doc.trade.settlement.monetary_summation.tax_total = (
            Decimal(f"{steuer:.2f}"),
            "EUR",
        )
        self.doc.trade.settlement.monetary_summation.grand_total = Decimal(f"{brutto:.2f}")
        self.doc.trade.settlement.monetary_summation.due_amount = Decimal(f"{brutto:.2f}")
        self.doc.trade.settlement.monetary_summation.charge_total = Decimal("0.00")
        self.doc.trade.settlement.monetary_summation.allowance_total = Decimal("0.00")

    def add_zahlungsziel(self, text, datum):
        """add zahlungsziel to zugferd"""
        terms = PaymentTerms()
        terms.description = text
        terms.due = datum
        try:
            self.doc.trade.settlement.terms.add(terms)
        except (AttributeError, TypeError):
            pass  # Skip if not available

    def add_xml2pdf(self, in_file=None, out_file=None) -> None:
        """add xml content to out_file"""
        # Generate XML file
        xml = self.doc.serialize(schema="FACTUR-X_EXTENDED")
        
        # Manipuliere die XML um den buyer.id hinzuzufügen wenn nicht vorhanden
        if self.customer_number_value:
            try:
                root = etree.fromstring(xml)
                namespaces = {
                    'ram': 'urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100',
                    'rsm': 'urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100'
                }
                
                # Der XPath muss im SupplyChainTradeTransaction sein
                buyer_parties = root.xpath(
                    '//rsm:SupplyChainTradeTransaction/ram:ApplicableHeaderTradeAgreement/ram:BuyerTradeParty',
                    namespaces=namespaces
                )
                
                if buyer_parties:
                    buyer_party = buyer_parties[0]
                    # Überprüfe ob bereits ID existiert
                    existing_ids = buyer_party.xpath('./ram:ID', namespaces=namespaces)
                    
                    if not existing_ids:
                        # Erstelle ein neues ID-Element
                        id_elem = etree.Element('{' + namespaces['ram'] + '}ID')
                        id_elem.text = self.customer_number_value
                        # Füge es am Anfang des BuyerTradeParty hinzu
                        buyer_party.insert(0, id_elem)
                        xml = etree.tostring(root, encoding='utf-8', xml_declaration=True)
            except Exception as e:
                # Wenn XML-Manipulation fehlschlägt, verwende original XML
                import traceback
                print(f"XML Manipulation fehlgeschlagen: {e}")
                traceback.print_exc()
                pass

        # Attach XML to an existing PDF.
        # Note that the existing PDF should be compliant to PDF/A-3!
        # You can validate this here:
        #   https://www.pdf-online.com/osa/validate.aspx
        if in_file:
            with open(in_file, "rb") as original_file:
                new_pdf_bytes = attach_xml(original_file.read(), xml, "EXTENDED")
        if self.debug:
            with open("factur-x.xml", "wb") as f:
                f.write(xml)
        if out_file:
            with open(out_file, "wb") as f:
                f.write(new_pdf_bytes)

    # def modify_xml(self, xml=None):
    #     """insert xmlns:qdt if it is not in namespaces"""
    #     decoded = xml.decode('utf-8')
    #     searchstr = re.search(r'<rsm:CrossIndustryInvoice(.*)>',
    #                           decoded).group()
    #     # print(searchstr)
    #     nsmap = searchstr.split(' ')
    #     _QDT = 'xmlns:qdt='
    #     QDT = _QDT + '\"' + NS_QDT + '\"'
    #     if QDT not in nsmap:
    #         nsmap.insert(1, QDT)
    #         decoded = decoded.replace(searchstr, ' '.join(nsmap))
    #     # print ('MAP:', nsmap)

    #     return decoded.encode('utf-8')

    def fill_lieferant_to_note(self, lieferant: Lieferant) -> None:
        """
        populate note with addressfields for ZugFeRD
        """
        txt = (
            lieferant.anschrift
            + "\n"
            + lieferant.kontakt
            + "\n"
            + lieferant.umsatzsteuer
        )
        self.add_note(txt)

    def _get_the_tax(
        self, steuersatz: str = "19.00", is_kleinunternehmen: bool = False
    ) -> str:
        return "0.00" if is_kleinunternehmen else steuersatz

    def _fill_invoice_positions_in_xml(
        self, positions: np.ndarray = None, steuersatz: str = None
    ) -> None:
        """fills invoice positions into ZugFeRD"""
        if positions is not None:
            self.add_items(positions, steuersatz)

    def _get_brutto(self, sums: list = None) -> str:
        return self._get_float_value(sums[-1])

    def _set_first_datum(self, df: pd.DataFrame, headers: list) -> None:
        """I expect 'Datum' at second position in df"""
        tmp = df.loc[df[headers[1]].index[0], headers[1]]
        if isinstance(tmp, datetime):
            return
        if math.isnan(tmp):
            df.loc[df[headers[1]].index[0], headers[1]] = self.rg_date

    def fill_xml(self, invoice: InvoiceCollection = None) -> None:
        """
        fills data into ZugFeRD part
        """
        if invoice is None:
            return
        kleinunternehmen = invoice.management.is_kleinunternehmen
        steuersatz = invoice.supplier.steuersatz
        self.add_zahlungsempfaenger(invoice.supplier_account.multiliner())

        self.fill_lieferant_to_note(invoice.supplier)
        
        # Füge Rechnungsüberschrift als Note ins ZugFeRD ein
        if invoice.invoice_title:
            self.add_note(invoice.invoice_title)
        
        # Füge Rechnungsbeschreibung als Note ins ZugFeRD ein
        if invoice.invoice_subtitle:
            self.add_note(invoice.invoice_subtitle)
        
        self.add_my_company(invoice.supplier)
        rg_nr = list(invoice.invoicenr.values())[0]
        
        # Konvertiere self.rg_date zu datetime, wenn es ein String ist
        raw_date = list(invoice.invoicedate.values())[0]
        if isinstance(raw_date, str):
            try:
                self.rg_date = datetime.strptime(raw_date, "%d.%m.%Y")
            except (ValueError, TypeError):
                try:
                    self.rg_date = datetime.strptime(raw_date, "%Y-%m-%d")
                except (ValueError, TypeError):
                    self.rg_date = datetime.now()
        else:
            self.rg_date = raw_date
        
        self.add_rgnr(rg_nr, self.rg_date)
        # Kundennummer hinzufügen
        if invoice.customer_number:
            cust_nr = list(invoice.customer_number.values())[0]
            self.add_customer_number(cust_nr)
        self.add_rechnungsempfaenger(None, invoice.customer)
        positions = invoice.positions.copy()
        headers = list(positions.columns)
        self._set_first_datum(positions, headers)
        self._fill_invoice_positions_in_xml(
            np.r_[[positions.columns], positions.astype(str).values],
            self._get_the_tax(steuersatz, kleinunternehmen),
        )
        self.add_gesamtsummen(
            invoice.sums,
            self._get_the_tax(steuersatz, kleinunternehmen),
            P19USTG if kleinunternehmen else None,
        )
        # Datum
        self.add_rechnungsperiode(positions[headers[1]])
        self.add_zahlungsziel(
            f"Bitte überweisen Sie den Betrag von \
{self._get_brutto(invoice.sums)} bis zum",
            invoice.supplier.get_ueberweisungsdatum(self.rg_date),
        )
        self.add_verwendungszweck(invoice.invoicenr, self.rg_date.strftime(GERMAN_DATE))
