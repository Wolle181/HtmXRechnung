"""
Module Invoice
"""

import pandas as pd
from src.lieferant import Lieferant
from src.kunde import Kunde
from src.konto import Konto
from src.steuerung import Steuerung


class Invoice(object):
    """Class Invoice handles collected data for invoice"""

    def __init__(self) -> None:
        self._customer: Kunde = None
        self._supplier: Lieferant = None
        self._supplier_account: Konto = None
        self._positions: pd.DataFrame = None
        self._invoicenr: dict = None
        self._customer_number: dict = None
        self._invoicedate: dict = None
        self._invoice_title: str = "Rechnung über Honorare"
        self._invoice_subtitle: str = "Für die nachfolgend aufgeführten Leistungen erlaube ich mir nach der Steuerberater\u00ADvergütungs\u00ADverordnung wie folgt abzurechnen:"
        self._sums: list = None
        self._management: Steuerung = None

    def __repr__(self) -> str:
        return f"Invoice: customer: '{repr(self.customer)}',\
 supplier: '{repr(self.supplier)}',\
 supplier_account: '{repr(self.supplier_account)}',\
 positions: '{self.positions}',\
 invoicenr: '{self.invoicenr}',\
 invoicedate: '{self.invoicedate}',\
 sums '{self.sums}'\
 management: '{repr(self.management)}'"

    @property
    def customer(self) -> Kunde:
        return self._customer

    @customer.setter
    def customer(self, value: Kunde) -> None:
        self._customer = value

    @property
    def supplier(self) -> Lieferant:
        return self._supplier

    @supplier.setter
    def supplier(self, value: Lieferant) -> None:
        self._supplier = value

    @property
    def supplier_account(self) -> Konto:
        return self._supplier_account

    @supplier_account.setter
    def supplier_account(self, value: Konto) -> None:
        self._supplier_account = value

    @property
    def positions(self) -> pd.DataFrame:
        return self._positions

    @positions.setter
    def positions(self, value: pd.DataFrame) -> None:
        self._positions = value

    @property
    def invoicenr(self) -> dict:
        return self._invoicenr

    @invoicenr.setter
    def invoicenr(self, value: dict) -> None:
        self._invoicenr = value

    @property
    def customer_number(self) -> dict:
        return self._customer_number

    @customer_number.setter
    def customer_number(self, value: dict) -> None:
        self._customer_number = value

    @property
    def invoicedate(self) -> dict:
        return self._invoicedate

    @invoicedate.setter
    def invoicedate(self, value: dict) -> None:
        self._invoicedate = value

    @property
    def invoice_title(self) -> str:
        return self._invoice_title

    @invoice_title.setter
    def invoice_title(self, value: str) -> None:
        self._invoice_title = value

    @property
    def invoice_subtitle(self) -> str:
        return self._invoice_subtitle

    @invoice_subtitle.setter
    def invoice_subtitle(self, value: str) -> None:
        self._invoice_subtitle = value

    @property
    def sums(self) -> list:
        return self._sums

    @sums.setter
    def sums(self, value: list) -> None:
        self._sums = value

    @property
    def management(self) -> Steuerung:
        return self._management

    @management.setter
    def management(self, value: Steuerung) -> None:
        self._management = value
