#!/usr/bin/env python3
"""
Test script to verify button visibility in Invoice Mapping Profiles tab
"""
import sys
from PyQt5.QtWidgets import QApplication
from DerivativeMill.derivativemill import DerivativeMill

app = QApplication(sys.argv)
window = DerivativeMill()

# Check if the Invoice Mapping Profiles tab exists
print(f"Tab widget exists: {hasattr(window, 'tabs')}")
print(f"Shipment map tab exists: {hasattr(window, 'tab_shipment_map')}")

# Manually click to tab 1 (Invoice Mapping Profiles) to trigger setup
print("\nSimulating tab switch to Invoice Mapping Profiles (index 1)...")
window.on_tab_changed(1)

# Check if buttons were created
if hasattr(window, 'tab_shipment_map'):
    buttons = window.tab_shipment_map.findChildren(type(window.process_btn).__bases__[0])
    print(f"Buttons found in tab_shipment_map: {len(buttons)}")
    for btn in buttons:
        print(f"  - {btn.text()}: visible={btn.isVisible()}, geometry={btn.geometry()}")

# Check if profile_combo_map was created
print(f"\nprofile_combo_map exists: {hasattr(window, 'profile_combo_map')}")
if hasattr(window, 'profile_combo_map'):
    print(f"profile_combo_map visible: {window.profile_combo_map.isVisible()}")
    print(f"profile_combo_map geometry: {window.profile_combo_map.geometry()}")

window.show()
print("\nWindow shown. Check if buttons are visible in Invoice Mapping Profiles tab.")
print("Look for buttons: 'Load Invoice File', 'Reset Current', 'Save Current Mapping As...', 'Delete Profile'")

sys.exit(app.exec_())
