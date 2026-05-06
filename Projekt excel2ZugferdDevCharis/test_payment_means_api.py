#!/usr/bin/env python3
"""Test PaymentMeans.information API"""

from drafthorse.models.payment import PaymentMeans

print("=" * 60)
print("TEST 1: CORRECT - Using .add() method")
print("=" * 60)
pm = PaymentMeans()
print(f"Initial type: {type(pm.information)}")
print(f"Initial children: {pm.information.children}")

pm.information.add("Line 1 of payment info")
pm.information.add("Line 2 of payment info")

print(f"After .add(): {type(pm.information)}")
print(f"Children: {pm.information.children}")
print("✓ SUCCESS!\n")

print("=" * 60)
print("TEST 2: WRONG - Direct string assignment")
print("=" * 60)
pm2 = PaymentMeans()
pm2.information = "Direct assignment"
print(f"After direct assignment: {type(pm2.information)}")
print(f"Value: {pm2.information}")

try:
    pm2.information.add("Try to add after direct assignment")
except AttributeError as e:
    print(f"✗ ERROR: {e}\n")

print("=" * 60)
print("CONCLUSION:")
print("=" * 60)
print("information is a StringContainer with .add() method")
print("Do NOT use direct assignment like: pm.information = 'text'")
print("ALWAYS use: pm.information.add('text')")
print("\nThe error likely means information was overwritten")
print("with a plain string somewhere in your code.")
