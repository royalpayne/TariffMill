#!/usr/bin/env python3
"""
Interactive OCR Template Creator

This script helps you create custom OCR extraction templates for your suppliers.
It guides you through analyzing a sample invoice and creating a template.
"""

import sys
from pathlib import Path
from ocr import SupplierTemplate, get_template_manager, preview_extraction

def print_header(text):
    """Print a formatted header"""
    print(f"\n{'='*70}")
    print(f"  {text}")
    print(f"{'='*70}")

def print_section(text):
    """Print a formatted section"""
    print(f"\n{text}")
    print("-" * 70)

def analyze_invoice(pdf_path):
    """Preview OCR text from invoice"""
    print_section("📄 Analyzing Invoice Text")

    try:
        preview = preview_extraction(pdf_path, max_lines=50)

        if not preview['is_scanned']:
            print("⚠️  This PDF appears to be digital (not scanned)")
            print("   It should be extracted with pdfplumber, not OCR")
            return None

        print(f"Is Scanned: ✓")
        print(f"Total lines: {preview['line_count']}")
        print(f"Total characters: {preview['char_count']}")
        print(f"\nExtracted Text:")
        print(preview['text_preview'])

        return preview['text_preview']

    except Exception as e:
        print(f"❌ Error analyzing invoice: {e}")
        return None

def get_supplier_name():
    """Get supplier name from user"""
    print_section("1️⃣  Enter Supplier Name")

    while True:
        name = input("   Supplier name (e.g., 'ACME Electronics'): ").strip()
        if name:
            return name
        print("   ❌ Please enter a supplier name")

def get_patterns():
    """Get regex patterns from user"""
    print_section("2️⃣  Configure Pattern Matching")

    print("""
Patterns help OCR find Part Numbers and Values in your invoice.

For example, if your invoice shows:
  Part Number: ABC-123     Price: $49.99

The patterns would match:
  - Part Number: ABC-123   (3-25 alphanumeric chars)
  - Price: $49.99          (dollar amount)
    """)

    # Part number pattern
    print("\nPart Number Pattern:")
    print("  Default: [A-Z0-9\\-_\\.]{3,25} (matches ABC-123, SKU12345, etc.)")
    part_pattern = input("  Enter pattern (or press Enter for default): ").strip()
    if not part_pattern:
        part_pattern = r'([A-Z0-9\-_\.]{3,25})'
    else:
        part_pattern = f'({part_pattern})'

    # Value pattern
    print("\nValue Pattern:")
    print("  Default: \\$?\\s*(\\d{1,10}(?:[,\\.]?\\d{1,3})*(?:\\.\\d{2})?)")
    print("  This matches: $49.99, 100.50, 1,234.56, etc.")
    value_pattern = input("  Enter pattern (or press Enter for default): ").strip()
    if not value_pattern:
        value_pattern = r'\$?\s*(\d{1,10}(?:[,\.]?\d{1,3})*(?:\.\d{2})?)'

    return {
        'part_number': part_pattern,
        'value': value_pattern
    }

def create_template(supplier_name, patterns):
    """Create and save the template"""
    print_section("3️⃣  Creating Template")

    try:
        template = SupplierTemplate(supplier_name)

        # Update patterns
        template.patterns['part_number_value'] = patterns['part_number']
        template.patterns['value_pattern'] = patterns['value']

        # Save template
        manager = get_template_manager()
        manager.save_template(template)

        print(f"✅ Template created successfully!")
        print(f"   Supplier: {supplier_name}")
        print(f"   Location: DerivativeMill/ocr/templates/{supplier_name}.json")

        return True

    except Exception as e:
        print(f"❌ Error creating template: {e}")
        return False

def test_template(pdf_path, supplier_name):
    """Test the template on the invoice"""
    print_section("4️⃣  Testing Template")

    try:
        from ocr import extract_from_scanned_invoice

        print(f"Testing extraction with supplier: {supplier_name}")
        df, metadata = extract_from_scanned_invoice(pdf_path, supplier_name=supplier_name)

        print(f"✅ Extraction successful!")
        print(f"   Rows extracted: {len(df)}")
        print(f"   Columns: {metadata['columns']}")
        print(f"\nFirst 10 rows:")
        print(df.head(10).to_string())

        return True

    except Exception as e:
        print(f"⚠️  Extraction had issues: {e}")
        print("\nThis may be expected if your patterns don't match perfectly.")
        print("You can edit the template JSON file to adjust patterns:")
        print(f"  DerivativeMill/ocr/templates/{supplier_name}.json")
        return False

def main():
    """Main interactive flow"""
    print_header("OCR Template Creator for DerivativeMill")

    # Get PDF file
    print_section("Select Your Invoice")
    pdf_path = input("  Enter path to scanned invoice PDF: ").strip()

    if not Path(pdf_path).exists():
        print(f"❌ File not found: {pdf_path}")
        sys.exit(1)

    # Analyze invoice
    print("\nAnalyzing your invoice...")
    text_preview = analyze_invoice(pdf_path)

    if text_preview is None:
        sys.exit(1)

    # Get supplier name
    supplier_name = get_supplier_name()

    # Get patterns
    patterns = get_patterns()

    # Create template
    if not create_template(supplier_name, patterns):
        sys.exit(1)

    # Test template
    test_success = test_template(pdf_path, supplier_name)

    # Final instructions
    print_header("✅ Template Created Successfully!")

    print(f"""
Your template is ready to use:

  Supplier: {supplier_name}
  File: DerivativeMill/ocr/templates/{supplier_name}.json

Next Steps:

  1. Test in the application:
     - Go to "Invoice Mapping Profiles" tab
     - Click "Load Invoice File"
     - Select your scanned PDF
     - Verify extracted data is correct

  2. To improve accuracy:
     - If extraction isn't perfect, edit the template JSON
     - Adjust regex patterns based on actual text
     - Test again with different invoices

  3. For multiple suppliers:
     - Run this script again for each supplier
     - Create separate templates for each

Questions? See OCR_TEMPLATE_SETUP_GUIDE.md for detailed help.
    """)

if __name__ == '__main__':
    main()
