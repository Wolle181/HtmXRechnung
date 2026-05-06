"""
Test für PaymentMeans in drafthorse
Zeigt die KORREKTE Syntax und Struktur
"""

from drafthorse.models.document import Document

def test_payment_means_complete():
    """Test der vollständigen PaymentMeans Struktur"""
    doc = Document()
    
    # 1. Type Code (Zahlungsart) setzen
    doc.trade.settlement.payment_means.type_code = "58"  # SEPA Überweisung
    print(f"✓ Type Code: {doc.trade.settlement.payment_means.type_code}")
    
    # 2. Information (Zahlungsinformation) hinzufügen
    doc.trade.settlement.payment_means.information.add("Zahlung per SEPA Überweisung")
    print(f"✓ Information: {doc.trade.settlement.payment_means.information}")
    
    # 3. Payee Account (Zahlungsempfänger Konto) - UNSER KONTO
    doc.trade.settlement.payment_means.payee_account.account_name = "Max Mustermann GmbH"
    doc.trade.settlement.payment_means.payee_account.iban = "DE89370400440532013000"
    doc.trade.settlement.payment_means.payee_account.proprietary_id = "123456789"
    print(f"✓ Payee Account Name: {doc.trade.settlement.payment_means.payee_account.account_name}")
    print(f"✓ Payee IBAN: {doc.trade.settlement.payment_means.payee_account.iban}")
    print(f"✓ Payee Proprietary ID: {doc.trade.settlement.payment_means.payee_account.proprietary_id}")
    
    # 4. Payee Institution (unsere Bank) - BIC
    doc.trade.settlement.payment_means.payee_institution.bic = "COBADEFF"
    print(f"✓ Payee BIC: {doc.trade.settlement.payment_means.payee_institution.bic}")
    
    # 5. Optional: Payer Account (Zahler/Kunde Konto)
    doc.trade.settlement.payment_means.payer_account.iban = "DE75512108001234567890"
    print(f"✓ Payer IBAN (optional): {doc.trade.settlement.payment_means.payer_account.iban}")
    
    print("\n✅ Alle Tests ERFOLGREICH!")
    return doc


def test_payment_means_with_text_parsing():
    """Test mit Text-Parsing wie in add_zahlungsempfaenger"""
    doc = Document()
    
    # Beispiel-Input Text
    text = "Max Mustermann GmbH\nIBAN: DE89370400440532013000\nBIC: COBADEFF"
    
    # Parsing und Setzen
    doc.trade.settlement.payment_means.type_code = "58"
    doc.trade.settlement.payment_means.information.add("Zahlung per SEPA Überweisung.")
    
    arr = text.split("\n")
    
    # 1. Kontoinhaber
    doc.trade.settlement.payment_means.payee_account.account_name = arr[0]
    print(f"✓ Account Name: {doc.trade.settlement.payment_means.payee_account.account_name}")
    
    # 2. IBAN (falls vorhanden)
    if len(arr) > 1:
        iban_value = arr[1].split(" ", 1)[1]
        doc.trade.settlement.payment_means.payee_account.iban = iban_value
        print(f"✓ IBAN: {doc.trade.settlement.payment_means.payee_account.iban}")
    
    # 3. BIC (falls genau 3 Zeilen)
    if len(arr) == 3:
        bic_value = arr[2].split(" ", 1)[1]
        doc.trade.settlement.payment_means.payee_institution.bic = bic_value
        print(f"✓ BIC: {doc.trade.settlement.payment_means.payee_institution.bic}")
    
    print("\n✅ Text-Parsing TEST ERFOLGREICH!")
    return doc


def test_full_structure():
    """Test der VOLLSTÄNDIGEN PaymentMeans Struktur"""
    doc = Document()
    
    pm = doc.trade.settlement.payment_means
    
    print("PaymentMeans Struktur:")
    print(f"  1. type_code: {pm.type_code} (StringField - erforderlich)")
    print(f"  2. information: {type(pm.information).__name__} (StringContainer)")
    print(f"  3. financial_card: {pm.financial_card} (FinancialCard)")
    print(f"  4. payer_account: {pm.payer_account} (PayerFinancialAccount)")
    print(f"  5. payee_account: {pm.payee_account} (PayeeFinancialAccount)")
    print(f"  6. payee_institution: {pm.payee_institution} (PayeeFinancialInstitution)")
    
    print("\nPayeeFinancialAccount Properties:")
    print(f"  - iban: {pm.payee_account.iban}")
    print(f"  - account_name: {pm.payee_account.account_name}")
    print(f"  - proprietary_id: {pm.payee_account.proprietary_id}")
    
    print("\nPayeeFinancialInstitution Properties:")
    print(f"  - bic: {pm.payee_institution.bic}")
    
    print("\n✅ STRUKTUR-TEST ERFOLGREICH!")


if __name__ == "__main__":
    print("=" * 60)
    print("TEST 1: Vollständige PaymentMeans Struktur")
    print("=" * 60)
    test_payment_means_complete()
    
    print("\n" + "=" * 60)
    print("TEST 2: PaymentMeans mit Text-Parsing")
    print("=" * 60)
    test_payment_means_with_text_parsing()
    
    print("\n" + "=" * 60)
    print("TEST 3: Vollständige Struktur-Übersicht")
    print("=" * 60)
    test_full_structure()
