# drafthorse PaymentMeans - Korrekte API Referenz

## Fehleranalyse
Der Fehler **"'Container' object has no attribute 'payee_account'"** deutet darauf hin, dass irgendwo in deinem Code `payment_means` als Container statt als `PaymentMeans`-Objekt behandelt wird. Dies ist wahrscheinlich ein Problem mit der Initialisierung oder einer fehlerhaften Umleitungslogik.

---

## 1. KORREKTE SYNTAX FÜR PAYMENTMEANS

### Struktur der PaymentMeans-Klasse
```python
from drafthorse.models.payment import PaymentMeans

# PaymentMeans hat diese Properties:
payment_means.type_code              # StringField (erforderlich)
payment_means.information            # StringContainer (MultiStringField)
payment_means.financial_card         # FinancialCard
payment_means.payer_account          # PayerFinancialAccount
payment_means.payee_account          # PayeeFinancialAccount ← FÜR KONTOINHABER/IBAN
payment_means.payee_institution      # PayeeFinancialInstitution ← FÜR BIC
```

---

## 2. KONTOINHABER (PAYEE ACCOUNT NAME) SETZEN

```python
# Kontoinhaber/Zahlungsempfänger Name
doc.trade.settlement.payment_means.payee_account.account_name = "Max Mustermann GmbH"
```

**Wichtig:** 
- Dies ist das **Kontoinhaber-Namentfeld** (BR-17 in ZUGFeRD)
- Es wird in `PayeeFinancialAccount` gespeichert, nicht direkt in PaymentMeans

---

## 3. IBAN SETZEN

```python
# IBAN (International Bank Account Number)
doc.trade.settlement.payment_means.payee_account.iban = "DE89370400440532013000"
```

**Wichtig:**
- Nur das IBAN-Nummern-Teil setzen (ohne "IBAN:" Präfix)
- Dies gehört zu `PayeeFinancialAccount` SubElement
- XML-Tag wird automatisch als `IBANID` generiert

---

## 4. BIC/SWIFT CODE SETZEN

```python
# BIC/Swift Code
doc.trade.settlement.payment_means.payee_institution.bic = "COBADEFF"
```

**Wichtig:**
- Dies gehört zu `PayeeFinancialInstitution` SubElement
- BIC muss 8 oder 11 Zeichen lang sein
- XML-Tag wird automatisch als `BICID` generiert

---

## 5. ZAHLUNGSART (TYPE_CODE)

```python
# Zahlungsart setzen
doc.trade.settlement.payment_means.type_code = "58"  # SEPA-Überweisung
# Andere Codes: "49" (Scheck), "ZZZ" (andere)
```

---

## 6. ZAHLUNGSINFORMATION (INFORMATION)

```python
# Zahlungsinformationen hinzufügen
doc.trade.settlement.payment_means.information.add("Zahlung per SEPA Überweisung")
```

**Wichtig:**
- `information` ist ein `StringContainer` (nicht a list!)
- Nutze `.add()` Methode, nicht direkte Zuweisung

---

## 7. VOLLSTÄNDIGES BEISPIEL

```python
from drafthorse.models.document import Document

doc = Document()

# Type Code setzen
doc.trade.settlement.payment_means.type_code = "58"  # SEPA

# Information hinzufügen
doc.trade.settlement.payment_means.information.add("SEPA-Überweisung")

# Kontoinhaber/Zahlungsempfänger
doc.trade.settlement.payment_means.payee_account.account_name = "Max Mustermann GmbH"

# IBAN
doc.trade.settlement.payment_means.payee_account.iban = "DE89370400440532013000"

# BIC
doc.trade.settlement.payment_means.payee_institution.bic = "COBADEFF"

# Optional: Kontonummer (proprietary)
doc.trade.settlement.payment_means.payee_account.proprietary_id = "123456789"
```

---

## 8. GESAMTSTRUKTUR VON PAYMENTMEANS

```
PaymentMeans
├── type_code: "58" (SEPA-Überweisung)
├── information: StringContainer
│   └── "Zahlung per SEPA Überweisung"
├── financial_card: FinancialCard
│   ├── card_holder_name
│   └── card_number
├── payer_account: PayerFinancialAccount (Zahler/Payer)
│   ├── iban
│   ├── account_name
│   └── proprietary_id
├── payee_account: PayeeFinancialAccount (Zahlungsempfänger) ← DEIN KONTO
│   ├── iban: "DE89370400440532013000"
│   ├── account_name: "Kontoinhaber Name"
│   └── proprietary_id: "123456789"
└── payee_institution: PayeeFinancialInstitution (Deine Bank)
    └── bic: "COBADEFF"
```

---

## 9. FEHLERSUCHE

### Fehler: "Container object has no attribute 'payee_account'"

**Ursachen:**
1. ❌ Falsch: `payment_means` ist ein Container-Objekt statt PaymentMeans
2. ❌ Falsch: Du bearbeitest das settlement.terms statt settlement.payment_means
3. ❌ Falsch: Du versuchst, auf `settlement.payee` zuzugreifen (das ist eine andere Klasse!)

**Lösungen:**
- ✅ Verwende NUR: `doc.trade.settlement.payment_means`
- ✅ NICHT: `doc.trade.settlement.payee` (das ist PayeeTradeParty, nicht PaymentMeans)
- ✅ NICHT: `doc.trade.settlement.terms` (das ist Container von PaymentTerms)

---

## 10. ANPASSUNG FÜR DEINEN CODE

Deine aktuelle Funktion `add_zahlungsempfaenger` sollte so aussehen:

```python
def add_zahlungsempfaenger(self, text):
    """set Zahlungsempfaenger zu korrektem Wert"""
    self.doc.trade.settlement.payment_means.type_code = "58"  # SEPA Überweisung
    
    # Information hinzufügen (nutze .add())
    self.doc.trade.settlement.payment_means.information.add(
        "Zahlung per SEPA Überweisung."
    )
    
    # Text parsen: "Name\nIBAN: DE...\nBIC: COBADEFF"
    arr = text.split("\n")
    
    # 1. Kontoinhaber (Payee Account Name)
    self.doc.trade.settlement.payment_means.payee_account.account_name = arr[0]
    
    # 2. IBAN (falls vorhanden)
    if len(arr) > 1:
        iban_parts = arr[1].split(" ", 1)
        self.doc.trade.settlement.payment_means.payee_account.iban = iban_parts[1]
    
    # 3. BIC (falls genau 3 Zeilen)
    if len(arr) == 3:
        bic_parts = arr[2].split(" ", 1)
        self.doc.trade.settlement.payment_means.payee_institution.bic = bic_parts[1]
```

---

## Zusammenfassung

| Was | Wo | Code |
|-----|----|----|
| Kontoinhaber | PayeeFinancialAccount | `.payee_account.account_name = "Name"` |
| IBAN | PayeeFinancialAccount | `.payee_account.iban = "DE..."` |
| BIC/SWIFT | PayeeFinancialInstitution | `.payee_institution.bic = "CODE"` |
| Zahlungsart | PaymentMeans | `.type_code = "58"` |
| Info-Text | PaymentMeans | `.information.add("Text")` |
