#!/usr/bin/env python3
"""
Create AROMATE Invoice Template - Correct Format

This creates a template for AROMATE invoices with the CORRECT format:

Invoice format:
  SKU# 1562485 76,080 PCS USD 0.6580 USD 50,060.64
  SKU# 2641486 15,120 PCS 0.7140 10,795.68
  SKU# 2641487 48,780 PCS 0.7140 34,828.92
  SKU# 2641488 15,840 PCS 0.7320 11,594.88

Pattern breakdown:
  SKU# XXXXXXX = SKU (part number)
  XXXXX PCS = Quantity
  USD X.XXXX = Unit price
  USD XXXXXXX.XX = Total price
"""

import re
import pandas as pd
import pdfplumber

def extract_invoice_data(pdf_path):
    """
    Extract SKU and price data from AROMATE invoice using regex

    Returns:
        DataFrame with columns: sku, quantity, unit_price, total_price
    """

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    # Pattern to match invoice lines:
    # Format 1: SKU# 1562485 76,080 PCS USD 0.6580 USD 50,060.64
    # Format 2: SKU# 2641486 15,120 PCS 0.7140 10,795.68 (no USD labels)

    # Single pattern that handles both formats
    # SKU# XXXXX QTY PCS [USD] UNIT_PRICE [USD] TOTAL_PRICE
    pattern = r'SKU#\s*(\d+)\s+(\d+(?:,\d{3})*)\s+PCS\s+(?:USD\s+)?([\d.]+)\s+(?:USD\s+)?([\d,]+\.\d{2})'

    matches = re.findall(pattern, text)

    if not matches:
        print("❌ Still no matches. Showing raw text for debugging:")
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if 'SKU#' in line:
                print(f"Line {i}: {line}")
        return None

    # Convert matches to DataFrame
    data = []
    for sku, qty, unit_price, total_price in matches:
        data.append({
            'part_number': sku,
            'quantity': int(qty.replace(',', '')),
            'unit_price': float(unit_price),
            'total_price': float(total_price.replace(',', ''))
        })

    df = pd.DataFrame(data)
    return df


# Main execution
if __name__ == '__main__':
    print("\n" + "="*70)
    print("AROMATE INVOICE DATA EXTRACTION")
    print("="*70)

    pdf_path = 'Input/CH_HFA001.pdf'

    print(f"\nExtracting from: {pdf_path}")
    print("-" * 70)

    df = extract_invoice_data(pdf_path)

    if df is not None and len(df) > 0:
        print(f"\n✅ Successfully extracted {len(df)} line items!")
        print("\nExtracted Data:")
        print(df.to_string(index=False))

        print("\n" + "="*70)
        print("EXTRACTED FIELDS:")
        print("="*70)
        print(f"Part Numbers (SKUs):")
        for sku in df['part_number']:
            print(f"  - {sku}")

        print(f"\nTotal Value: USD {df['total_price'].sum():,.2f}")
        print(f"Total Quantity: {df['quantity'].sum():,} PCS")

        # Save to CSV for verification
        output_file = 'aromate_invoice_extracted.csv'
        df.to_csv(output_file, index=False)
        print(f"\n💾 Saved to: {output_file}")

    else:
        print("❌ Could not extract data")
        print("\nDEBUGGING INFO:")
        print("Check if CH_HFA001.pdf is the correct invoice file")
        print("Expected format: SKU# XXXXXX XXXXX PCS ...")
