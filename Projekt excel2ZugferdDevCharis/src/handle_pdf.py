"""
Module handle_pdf
"""

# -*- coding: utf8 -*-

from datetime import datetime
import os
import sys
from fpdf import FPDF
from fpdf.fonts import FontFace
from fpdf.enums import TableCellFillMode, OutputIntentSubType
from fpdf.output import PDFICCProfile

# import src.handle_zugferd as handle_zugferd
import src.handle_girocode as gc
import pandas as pd
import numpy as np
from src.invoice_collection import InvoiceCollection
import math
import locale
import decimal
from src.constants import P19USTG, GERMAN_DATE, GERMAN_DATE_SHORT

LEFTofABSENDER = 135


class PDF(FPDF):
    """
    Base Class for my PDF
    """

    def __init__(self, footer_txt: str = ""):
        super().__init__()
        self.footer_txt = footer_txt
        self.table_lines = None
        self.table_head_color = None
        self.table_head = None
        self.table_fill_color = None
        self.table_widths = None
        self.sum_table_widths = None

    def footer(self) -> None:
        """
        declare a footer for the pdf file
        """
        if len(self.footer_txt) == 0:
            return
        self.set_y(-15)
        self.set_font_size(size=8)
        self.cell(
            0,
            1,
            self.footer_txt,
            align="C",
        )
        self.ln()
        self.set_font(None, "I", 8)
        self.cell(0, 10, f"Seite {self.page_no()}/{{nb}}", align="C")

    def header(self) -> None:
        if not hasattr(self, "title"):
            return
        # Setting font: origin font bold 15
        self.set_font(None, "B", 15)
        # Calculating width of title and setting cursor position:
        width = self.get_string_width(self.title) + 6
        # self.set_x((210 - width) / 2)
        self.set_x(25)
        # Setting colors for frame, background and text:
        # self.set_draw_color(0, 80, 180)
        # self.set_fill_color(230, 230, 0)
        # self.set_text_color(220, 50, 50)
        # Setting thickness of the frame (1 mm)
        self.set_line_width(1)
        # Printing title:
        self.cell(
            width,
            9,
            self.title,
            # border=1,
            new_x="LMARGIN",
            new_y="NEXT",
            align="L",
            # fill=True,
        )
        # Performing a line break:
        self.ln()

    def print_faltmarken(self) -> None:
        """
        prints Falt- und Lochmarkierungen
        """
        self.line(4, 105, 8, 105)
        self.line(3, 148.5, 6, 148.5)  # Lochmarke
        self.line(4, 210, 8, 210)

    def _set_section(
        self, section: str, x: float, y: float, style: str, size: int
    ) -> None:
        """sets section, xy position, font"""
        if section is not None:
            self.start_section(section)
        self.set_xy(x, y)
        self.set_font(None, style, size)

    def print_absender(self, adress: str) -> None:
        """
        prints Absenderdaten
        """
        self._set_section("Absender", LEFTofABSENDER, 25.5, "", 11)
        arr = adress.splitlines()
        self.multi_cell(0, 5, "\n".join(arr[1:]))
        self._set_section("Abs-kurz", 25, 60.5, "U", 6)
        abs_kurz = ", ".join([arr[0], arr[-2], arr[-1]])
        self.cell(105, 1, abs_kurz)
        self.ln()

    def print_bundesland(self, bundesland: str) -> None:
        """
        prints Bundesland
        """
        self._set_section("Bundesland", LEFTofABSENDER, 54.5, "", 10)
        self.cell(105, 1, bundesland)

    def print_leistungszeitraum(self, von: str, bis: str) -> None:
        """
        prints Rechnungszeitraum
        """
        self._set_section("Leistungszeitraum", 25, 105, "", 10)
        self.cell(105, 1, f"Leistungszeitraum: {von} - {bis}")

    def print_kontakt(self, daten: str) -> None:
        """
        prints Kontakt
        """
        self._set_section("Kontakt", LEFTofABSENDER, 60, "", 10)
        self.multi_cell(105, 5, daten)

    def print_adress(self, adress: str) -> None:
        """
        prints Empfänger
        """
        self._set_section("Empfänger", 25, 63.6, "", 12)
        self.multi_cell(105, 5, adress)
        self.ln()

    def print_bezug(self, text: str) -> None:
        """
        prints Betreffzeile (mit Unterstützung für Zeilenumbrüche)
        """
        self._set_section("Betreff", 25, 97.4, "", 10)
        self.multi_cell(0, 5, text)
        self.ln()

    def print_invoice_title(self) -> None:
        """
        prints Invoice Title (Rechnungsüberschrift) über der Tabelle
        Rechtsbündig in Spalte C mit Schriftgröße 12 (wie Rechnungs-Nr.)
        """
        title = self.invoice.invoice_title if self.invoice else "Rechnung über Honorare"
        # Spalte C: von 52mm bis 121mm (69mm breit)
        # Rechtsbündig positionieren
        self.set_xy(52, self.get_y())
        self.set_font("dejavu-sans", "B", 12)
        self.cell(69, 5, title, align="R")  # 69mm breit, rechtsbündig
        self.ln(7)

    def print_invoice_subtitle(self) -> None:
        """
        prints Invoice Subtitle (Beschreibungstext) ab Spalte A
        Mit Zeilenumbruch und kleinerer Schriftgröße
        """
        subtitle = self.invoice.invoice_subtitle if self.invoice else ""
        if subtitle:
            self.ln(3)  # Abstand nach oben
            # Position: Ab Spalte A (25mm Margin)
            self.set_xy(25, self.get_y())
            self.set_font("dejavu-sans", "", 10)
            self.multi_cell(140, 4, subtitle, align="L")  # Linksbündig, ohne Blocksatz
            self.ln(8)

    def _printTableHeader(self):
        """set table Header"""
        y = max(135, self.get_y() + 3)
        self._set_section("Rechnungspositionen", 25, y, "", 10)
        if self.table_lines:
            self.set_draw_color(self.table_lines)
        else:
            self.set_draw_color(0, 0, 255)
        self.set_line_width(0.3)

    def _getHeadingsStyle(self):
        """return headings_style"""
        return FontFace(
            emphasis="BOLD",
            color=self.table_head_color if self.table_head_color else 255,
            fill_color=self.table_head if self.table_head else (255, 100, 0),
        )

    def _getCellFillColor(self):
        return self.table_fill_color if self.table_fill_color else (244, 235, 255)

    def _set_variable_Breite(self, arr: list) -> list:
        fixeBreite = sum(arr)
        variableBreite = self._getTableWidth() - fixeBreite
        if variableBreite > 1:
            arr.insert(2, variableBreite)
            return arr
        return None

    def _calcColWidths(self, lengths: list) -> list:
        if lengths:
            arr = []
            for index, val in enumerate(lengths):
                # index 2 ist die Spalte
                # Bezeichnung, die ich als
                # eine variable Spalte haben möchte
                if index != 2:
                    arr.append(math.ceil(val * (2.4 if val > 3 else 2.9)))
        return self._set_variable_Breite(arr)

    def _getColWidths(self, lengths: list) -> tuple:
        # 6 Spalten: Pos, Datum, Beschreibung, Anzahl/Gebührensatz, Preis, Summe
        # Einheit (ehemals 15mm) wird auf Beschreibungsspalte aufgeteilt → 165mm gesamt
        return (11, 15, 77, 19, 24, 19)

    def _getTableWidth(self) -> int:
        return 165
        # return (sum(self.table_widths) if self.table_widths
        #         else 165)

    def _format_table_headers(self, arr: list) -> None:
        """Format header row: split long words, add 'in €' to Wert/Gebühr"""
        header = arr[0]
        # Spalte D (Index 3): Gebührensatz / Arbeitnehmer umbrechen
        col_d = str(header[3]).replace("\n", "").replace("-", "").strip()
        if "Gebührensatz" in col_d:
            header[3] = "Gebühren-\nsatz"
        elif "Arbeitnehmer" in col_d:
            header[3] = "Arbeit-\nnehmer"
        # Spalte F (Index 4 nach Entfernung von Einheit) und G (Index 5): "in €"
        header[4] = header[4] + "\nin €"
        header[5] = header[5] + "\nin €"

    def print_positions(self, arr: list, lengths: list = None) -> None:
        """
        prints Table with Positions
        """
        # Alle Zeilen auf exakt 7 Spalten kürzen (A-G),
        # dann Spalte E (Index 4, Einheit) für PDF entfernen —
        # sie wird nur im ZUGFeRD-XML benötigt.
        # arr kann ein numpy-Array sein → erst in Liste von Listen umwandeln
        arr = [list(row)[:7] for row in arr]
        for row in arr:
            del row[4]  # Einheit (Stück/Std. etc.) nicht im PDF anzeigen

        self._format_table_headers(arr)
        self._printTableHeader()
        last = None

        # Berechne dynamische line_height basierend auf Zeilenumbrüchen
        max_lines = 1
        for data_row in arr[1:]:  # Skip header
            for cell in data_row:
                if isinstance(cell, str) and '\n' in cell:
                    lines = cell.count('\n') + 1
                    max_lines = max(max_lines, lines)

        # Basis line_height + extra für jede zusätzliche Zeile (reduzierte Abstände)
        dynamic_line_height = 5 + (max_lines - 1) * 3.5

        # Schriftgröße vor Tabelle setzen - 1 Punkt kleiner
        self.set_font_size(9)

        text_align = (
            "CENTER", "LEFT", "LEFT", "CENTER", "RIGHT", "RIGHT"
        )  # 6 Spalten: Einheit (ehemals Index 4) wird im PDF nicht angezeigt

        col_widths = self._getColWidths(lengths)

        with self.table(
            borders_layout="NO_HORIZONTAL_LINES",
            cell_fill_color=self._getCellFillColor(),
            cell_fill_mode=TableCellFillMode.ROWS,
            col_widths=col_widths,
            text_align=text_align,
            align="RIGHT",
            headings_style=self._getHeadingsStyle(),
            line_height=dynamic_line_height,
            width=self._getTableWidth(),
            padding=(1, 0, 1, 0),
            v_align="TOP",
        ) as table:
            for data_row in arr:
                if data_row[1] == "":  # Datum
                    data_row[1] = last
                last = data_row[1]
                table.row(data_row)

    def _getSummenStartX(self) -> int:
        """x-Position des Summenblocks = linker Rand + A + B + C"""
        col_widths = self._getColWidths(None)
        return 25 + col_widths[0] + col_widths[1] + col_widths[2]

    def _printSummenHeader(self) -> None:
        self.set_xy(self._getSummenStartX(), self.get_y())
        self.set_font(None, "", size=9)
        if self.table_lines:
            self.set_draw_color(self.table_lines)
        else:
            self.set_draw_color(255, 0, 255)
        self.set_line_width(0.3)

    def _getTableFillColor(self) -> tuple:
        return (
            self.table_fill_color
            if hasattr(self, "table_fill_color")
            else (244, 235, 255)
        )

    def _getSummenColWidths(self) -> tuple:
        if self.sum_table_widths:
            return self.sum_table_widths
        # Bezeichnung = D + F (ohne Einheit), Betrag = G
        col_widths = self._getColWidths(None)
        return (col_widths[3] + col_widths[4], col_widths[5])

    def _getSummenTableWidth(self) -> int:
        if self.sum_table_widths:
            return sum(self.sum_table_widths)
        return sum(self._getSummenColWidths())

    def print_summen(self, arr: list) -> None:
        """
        print Summen - halten die Summen, Umsatzsteuer und Brutto IMMER zusammen!
        Die 3 Zeilen dürfen NIEMALS aufgeteilt werden!
        """
        # 1-2 Zeilen Abstand zur vorherigen Tabelle
        self.ln(5)
        
        # print(arr)
        # Realistische Höhe für 3 Zeilen Summen (line_height=4.5 + padding)
        summen_height = 25  # mm - für 3 Zeilen mit Abstand
        
        # Überprüfe genau, ob genug Platz vorhanden ist
        current_y = self.get_y()
        page_height = self.h - 25  # Minus Footer und Sicherheitsabstand
        available_space = page_height - current_y
        
        # Wenn nicht mindestens 25mm Platz, sofort neue Seite
        if available_space < summen_height:
            self.add_page()
        
        self._printSummenHeader()
        # headings_style = FontFace(emphasis="BOLD", color=255,
        # fill_color=self.table_head
        # if hasattr(self, "table_head")
        # else (255, 100, 0))
        with self.table(
            borders_layout="NO_HORIZONTAL_LINES",
            cell_fill_color=self._getTableFillColor(),
            cell_fill_mode=TableCellFillMode.ROWS,
            col_widths=self._getSummenColWidths(),
            text_align=(
                "RIGHT",
                "RIGHT",
            ),
            first_row_as_headings=False,
            align="RIGHT",
            # headings_style=headings_style,
            line_height=4.5,
            width=self._getSummenTableWidth(),
            padding=(0.5, 0, 0.5, 0),
            v_align="TOP",
        ) as table:
            for data_row in arr:
                table.row(data_row)
        # Position zurücksetzen für korrekte Fußzeilen-Zentrier
        self.set_x(0)

    def print_kleinunternehmerregelung(self, grund: str) -> None:
        """
        prints Begründung Kleinunternehmerregelung
        """
        self.start_section("Kleinunternehmen", 0)
        self.cell(105, 1, grund)
        self.ln(2)

    def print_abspann(self, text: str) -> None:
        """
        print Abspann
        """
        self.ln()
        self.ln()
        self.start_section("Abspann", 0)
        self.set_font_size(10)
        self.multi_cell(0, None, text)
        self.ln()

    def print_signature(self, sig_path: str = None) -> None:
        """
        print Unterschrift zwischen Abspann und Name
        """
        if sig_path is None:
            from pathlib import Path
            if getattr(sys, 'frozen', False):
                base = Path(sys.executable).parent
            else:
                base = Path(__file__).parent.parent
            # Suche: neben EXE → APPDATA → assets (Dev)
            for candidate in [
                base / "signatur.png",
                Path(os.getenv("APPDATA", "")) / "excel2zugferd" / "signatur.png",
                base / "assets" / "signatur.png",
            ]:
                if candidate.exists():
                    sig_path = str(candidate)
                    break
        
        # Prüfe ob Datei existiert
        if not sig_path or not os.path.exists(sig_path):
            return
        
        # Unterschrift einfügen: 40mm breit, proportionale Höhe
        try:
            self.image(sig_path, x=25, w=40)
            self.ln(2)
        except Exception as e:
            print(f"Unterschrift konnte nicht eingefügt werden: {e}")

    def print_logo(self, fn: str) -> None:
        """
        print Logo
        """
        if fn is None:
            return
        self.start_section("Logo", 0)
        rect1 = 27, 30, 20, 20
        self.set_draw_color(255)
        self.rect(*rect1)
        self.image(fn, *rect1, keep_aspect_ratio=True)
        self.set_draw_color(0)

    def print_qrcode(self, img: object) -> None:
        """
        print qrcode
        """
        if img is None:
            return
        # rect1 = 27, 30, 20, 20
        self.start_section("GiroCode", 0)
        self.set_draw_color(255)
        # self.rect(*rect1)
        self.image(img.get_image(), w=30, h=30, keep_aspect_ratio=True)
        self.set_draw_color(0)
        self.set_x(28)
        self.set_font(None, "B", size=10)
        self.cell(80, 1, "Bezahlen via GiroCode")
        self.set_font(None, "", size=10)

    def uniquify(self, path: str, appendix: str = None) -> str:
        """
        make unique Path from filename
        """
        fn, ext = os.path.splitext(path)
        counter = 1

        if appendix is not None:
            equal = False
            for i in range(len(appendix) - 1):
                equal |= appendix[-(i + 1)] == fn[-(i + 1)]
            fn = f"{fn}{appendix}" if not equal else fn
            path = f"{fn}" + ext

        while os.path.exists(path):
            path = f"{fn} ({str(counter)})" + ext
            counter += 1
        return path


class Pdf(PDF):
    """
    Klasse Pdf
    """

    def __init__(self, logo_fn=None) -> None:
        super().__init__()
        self.logo_fn = logo_fn
        self.qrcode_img = None
        self.invoice: InvoiceCollection = None
        self.set_fonts_and_other_stuff()
        locale.setlocale(locale.LC_ALL, "de_DE.UTF-8")

    def set_fonts_and_other_stuff(self) -> None:
        """
        import and embed TTF Font to use € in text
        """
        # Bestimme den Basis-Pfad für die Ressourcen
        # Bei PyInstaller: sys._MEIPASS ist bereits das _internal Verzeichnis
        # Sonst: das Hauptprojektverzeichnis
        if getattr(sys, 'frozen', False):
            # Bei PyInstaller ist sys._MEIPASS das _internal Verzeichnis
            fonts_path = os.path.join(sys._MEIPASS, "Fonts")
            icc_profile_path = os.path.join(sys._MEIPASS, "sRGB2014.icc")
        else:
            # Bei normalem Python: _internal Verzeichnis
            project_base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            fonts_path = os.path.join(project_base, "_internal", "Fonts")
            icc_profile_path = os.path.join(project_base, "_internal", "sRGB2014.icc")
        
        self.add_font(
            "dejavu-sans", style="", fname=os.path.join(fonts_path, "DejaVuSansCondensed.ttf")
        )
        self.add_font(
            "dejavu-sans",
            style="b",
            fname=os.path.join(fonts_path, "DejaVuSansCondensed-Bold.ttf"),
        )
        self.add_font(
            "dejavu-sans",
            style="i",
            fname=os.path.join(fonts_path, "DejaVuSansCondensed-Oblique.ttf"),
        )
        self.add_font(
            "dejavu-sans",
            style="bi",
            fname=os.path.join(fonts_path, "DejaVuSansCondensed-BoldOblique.ttf"),
        )
        # use the font imported
        self.set_font("dejavu-sans")
        self.set_lang("de_DE")
        # set left, top and right margin for document
        self.set_margins(25, 16.9, 20)
        
        # ICC Profile - optional (nicht kritisch, kann fehlen)
        if os.path.exists(icc_profile_path):
            try:
                with open(icc_profile_path, "rb") as iccp_file:
                    icc_profile = PDFICCProfile(
                        contents=iccp_file.read(), n=3, alternate="DeviceRGB"
                    )
                    self.add_output_intent(
                        OutputIntentSubType.PDFA,
                        "sRGB",
                        "IEC 61966-2-1:1999",
                        "http://www.color.org",
                        icc_profile,
                        "sRGB2014 (v2)",
                    )
            except Exception:
                # ICC-Profil nicht verfügbar, weitermachen ohne
                pass

    def fill_header(self) -> None:
        """
        populate header with data
        """
        self.set_title(self.invoice.supplier.betriebsbezeichnung)
        self.footer_txt = self.invoice.supplier_account.oneliner()
        self.set_author(self.invoice.supplier.name)

        self.table_head = (
            255  # white fill-color of Table Header (30, 144, 255)
            # DodgerBlue1
        )
        self.table_head_color = 10  # 0  # black text-color of Table Header
        self.table_lines = 0  # black (0, 0, 255)  # Blue
        self.table_fill_color = 220  # lightgrey
        self.table_widths = (10, 21, 61, 16, 15, 21, 21)
        # (11, 22, 61, 16, 20, 21, 21)
        self.table_lines = 120  # darkgrey

        self.add_page()
        self.print_faltmarken()
        self.print_logo(self.logo_fn)
        self.print_absender(self.invoice.supplier.anschrift)
        bundesland = self.invoice.supplier.bundesland
        if bundesland and len(bundesland) > 0:
            self.print_bundesland(bundesland)
        self.print_kontakt(
            self.invoice.supplier.kontakt + "\n\n" + self.invoice.supplier.umsatzsteuer
        )

    def _fill_girocode(self, brutto, rg_nr, datum):
        girocode = gc.Handle_Girocode(
            self.invoice.supplier_account.bic,
            self.invoice.supplier_account.iban,
            self.invoice.supplier_account.name,
        )
        self.qrcode_img = girocode.girocodegen(
            brutto, f"{list(rg_nr.keys())[0]} {list(rg_nr.values())[0]} vom {datum}"
        )
        self.print_qrcode(self.qrcode_img)

    def _fill_kleinunternehmen(self) -> None:
        if self.invoice.management.is_kleinunternehmen:
            self.print_kleinunternehmerregelung(P19USTG)

    def get_maxlengths(self, df: pd.DataFrame) -> list:
        """
        get array with max string length of each column in pandas DataFrame
        """
        lenArr = []
        # print("get_maxlengths:\n", df)
        for c in df:
            theLength = max(df[c].astype(str).map(len).max(), len(c))
            # print('Max length of column %s: %s' % (c, theLength))
            lenArr.append(theLength)
        return lenArr

    def _currency(self, amount: float | decimal.Decimal) -> str:
        """return str as locale currency"""
        return locale.currency(amount, grouping=True)

    def _currency_no_symbol(self, amount: float | decimal.Decimal) -> str:
        """return str as locale currency without symbol"""
        return locale.currency(amount, symbol=False, grouping=True)

    def get_von_bis_dates(self) -> list:
        """get min and max dates from positions

        Returns:
            list: min and max as list
        """
        df = self.invoice.positions
        retval = df.copy()
        headers = list(df.columns)
        # Convert to datetime, invalid parsing will be set as NaT
        retval[headers[1]] = pd.to_datetime(
            retval[headers[1]], errors="coerce"  # Datum
        )
        # Drop NaT values
        valid_dates = retval[headers[1]].dropna()
        if valid_dates.empty:
            return [None, None]
        von = valid_dates.min()
        bis = valid_dates.max()
        return [von, bis]

    def _set_first_datum(self, df: pd.DataFrame, headers: list) -> None:
        """I expect Datum at second position in df"""
        if df.loc[df[headers[1]].index[0], headers[1]] == "":
            df.loc[df[headers[1]].index[0], headers[1]] = list(
                self.invoice.invoicedate.values()
            )[0].strftime(GERMAN_DATE_SHORT)

    def _change_values_to_german(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        replace column Datum with german date %d.%m.%Y and columns Anzahl,
        Preis and Summe as float values with Komma separated
        """
        retval = df.copy()
        headers = list(df.columns)
        # retval.style.format({datum: lambda t: t.strftime("%d.%m.%Y")
        #                      if len(t) > 0 else ""})
        retval[headers[1]] = pd.to_datetime(
            retval[headers[1]], errors="coerce"  # Datum
        ).dt.strftime(GERMAN_DATE_SHORT)
        # substitute NaN by ""
        retval[headers[1]] = retval[headers[1]].fillna("")
        self._set_first_datum(retval, headers)
        # print("_change_values_to_german:\n", retval)
        retval[headers[3]] = retval[headers[3]].replace("", np.nan).fillna(1).apply(
            lambda x: x if isinstance(x, str) and "/" in x else "{:n}".format(x)
        )  # Anzahl/Gebührensatz: Brüche (z.B. "2/10") als Text beibehalten
        retval[headers[5]] = retval[headers[5]].apply(  # Preis
            lambda x: self._currency_no_symbol(x)
        )
        retval[headers[6]] = retval[headers[6]].apply(  # Summe
            lambda x: self._currency_no_symbol(x)
        )
        return retval

    def get_invoice_positions(self, positions: pd.DataFrame = None) -> dict:
        """
        return dict with array of array of positions for invoice
        and maxlengths of columns
        {'daten': np.r_[], 'maxlengths': []}
        """
        retval = self._change_values_to_german(positions)
        # print(retval)
        lenArr = self.get_maxlengths(retval)
        # print(lenArr)

        # return {'daten': np.r_[line.values, retval.astype(str).values],
        return {
            "daten": np.r_[[retval.columns], retval.astype(str).values],
            "maxlengths": lenArr,
        }

    def _fill_positions(self) -> None:
        # print(self.split_dataframe_by_SearchValue(AN, "Pos."))
        theDict = self.get_invoice_positions(self.invoice.positions)
        if len(theDict) > 0:
            self.print_positions(theDict["daten"], theDict["maxlengths"])

    def _fill_abspann(self, brutto: str, ueberweisungsdatum: datetime) -> None:
        # Zahlungstext
        zahlungstext = f"Bitte überweisen Sie den Betrag von {brutto} bis zum \
{ueberweisungsdatum.strftime(GERMAN_DATE)} auf \
u.a. Konto."
        
        # Abspann (Mit freundlichen Grüßen + Name)
        abspann_greeting = (
            self.invoice.management.abspann
            if self.invoice.management.abspann
            and len(self.invoice.management.abspann) > 1
            else (
                "Mit freundlichen Grüßen"
                if self.invoice.supplier.name
                else ""
            )
        )
        
        supplier_name = self.invoice.supplier.name if self.invoice.supplier.name else ""
        
        # Ausgabe: Zahlungstext
        self.print_abspann(zahlungstext)
        
        # Überprüfe, ob der Grüße+Unterschrift+Name Block zusammen passt (ca. 35mm nötig)
        signature_block_height = 35  # mm - Platz für Grüße + Unterschrift + Name
        current_y = self.get_y()
        page_height = self.h - 25  # Minus Footer
        available_space = page_height - current_y
        
        # Falls nicht genug Platz, neue Seite
        if available_space < signature_block_height:
            self.add_page()
        
        # Abspann mit Unterschrift
        self.ln()
        self.ln()
        self.start_section("Gruss", 0)
        self.set_font_size(10)
        self.multi_cell(0, None, abspann_greeting)
        self.ln(1)
        
        # Unterschrift einfügen
        self.print_signature()
        
        # Name
        if supplier_name:
            self.set_font_size(10)
            self.multi_cell(0, None, supplier_name)
            self.ln()

    def _get_value(self, tuple) -> str:
        _, v = tuple
        return v

    def _toLocaleFloatStr(self, inp: str) -> str:
        """convert string witch float to locale float string"""
        return locale.format_string("%.f", float(inp), grouping=True)

    def get_invoice_sums(self):
        """return array of invoice sums"""
        netto = self._get_value(self.invoice.sums[0])
        umsatzsteuer = self._get_value(self.invoice.sums[1])
        brutto = self._get_value(self.invoice.sums[-1])
        satz = self._toLocaleFloatStr(self.invoice.supplier.steuersatz)
        UST = f"zzgl. Umsatzsteuer {satz}%:"
        return [
            ("Summe netto:", self._currency(netto)),
            (UST, self._currency(umsatzsteuer)),
            ("Bruttobetrag:", self._currency(brutto)),
        ]

    def fill_pdf(self, invoice: InvoiceCollection) -> None:
        """
        set own data
        """
        self.invoice = invoice
        self.fill_header()

        self.print_adress(self.invoice.customer.anschrift)
        rg_nr = self.invoice.invoicenr  # _get_rg_nr()
        # today = datetime.now()
        datum = list(invoice.invoicedate.values())[0]
        
        # Stelle sicher, dass datum ein datetime ist, nicht ein String
        if isinstance(datum, str):
            try:
                datum = datetime.strptime(datum, "%d.%m.%Y")
            except (ValueError, TypeError):
                try:
                    datum = datetime.strptime(datum, "%Y-%m-%d")
                except (ValueError, TypeError):
                    datum = datetime.now()
        
        von, bis = self.get_von_bis_dates()
        ueberweisungsdatum = self.invoice.supplier.get_ueberweisungsdatum(datum)

        # Rechnungsnummer + Kundennummer kombiniert ausgeben
        rg_text = f"{list(rg_nr.keys())[0]} {list(rg_nr.values())[0]} vom {datum.strftime(GERMAN_DATE)}"
        
        # Kundennummer falls vorhanden hinzufügen
        if self.invoice.customer_number:
            cust_nr = list(self.invoice.customer_number.values())[0]
            if cust_nr and str(cust_nr).strip().upper() != "NONE":
                rg_text += f"\nKunden-Nr: {cust_nr}"
        
        # Leistungszeitraum hinzufügen (wenn beide Daten vorhanden und verschieden)
        if von and bis and von != bis:
            rg_text += f"\nLeistungszeitraum: {von.strftime(GERMAN_DATE)} - {bis.strftime(GERMAN_DATE)}"
        
        self.print_bezug(rg_text)
        
        # Überschrift "Rechnung über Honorare" ausgeben
        self.print_invoice_title()
        
        # Beschreibungstext ausgeben
        self.print_invoice_subtitle()

        self._fill_positions()

        summen = self.get_invoice_sums()
        brutto = self._get_value(summen[-1])
        self.print_summen(summen)

        self._fill_kleinunternehmen()
        # if self.lieferantensteuerung.create_xml:
        #     self.fill_xml(rg_nr, summen, brutto, ueberweisungsdatum, datum)
        self._fill_abspann(brutto, ueberweisungsdatum)

        if self.invoice.management.create_girocode:
            self._fill_girocode(
                locale.atof(brutto.strip(" €")), rg_nr, datum.strftime(GERMAN_DATE)
            )
