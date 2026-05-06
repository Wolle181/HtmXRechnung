"""Test PaymentMeans structure in drafthorse"""
from drafthorse.models.document import Document
from drafthorse.models.payment import PaymentMeans

# Test the PaymentMeans structure
doc = Document()
pm = doc.trade.settlement.payment_means

# Try accessing payee_account
print("1. Checking payment_means attributes:")
print(f"   type: {type(pm)}")
print(f"   has payee_account: {hasattr(pm, 'payee_account')}")
print(f"   payee_account value: {pm.payee_account}")
print(f"   payee_account type: {type(pm.payee_account)}")

# Try setting values
print("\n2. Setting values:")
pm.type_code = "58"
pm.payee_account.account_name = "Kontoinhaber"
pm.payee_account.iban = "DE89370400440532013000"
pm.payee_institution.bic = "COBADEFF"

print(f"   account_name: {pm.payee_account.account_name}")
print(f"   iban: {pm.payee_account.iban}")
print(f"   bic: {pm.payee_institution.bic}")

print("\n3. PaymentMeans structure recap:")
print(f"   type_code: {pm.type_code}")
print(f"   information: {pm.information}")
print(f"   financial_card: {pm.financial_card}")
print(f"   payer_account: {pm.payer_account}")
print(f"   payee_account: {pm.payee_account}")
print(f"   payee_institution: {pm.payee_institution}")

print("\n4. PayeeFinancialAccount structure:")
print(f"   iban: {pm.payee_account.iban}")
print(f"   account_name: {pm.payee_account.account_name}")
print(f"   proprietary_id: {pm.payee_account.proprietary_id}")

print("\n5. PayeeFinancialInstitution structure:")
print(f"   bic: {pm.payee_institution.bic}")
