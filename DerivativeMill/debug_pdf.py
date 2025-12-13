#!/usr/bin/env python3
"""
PDF Diagnostic Tool

Helps debug why PDF extraction fails by showing:
- PDF structure (pages, content type)
- Available tables on each page
- Text extraction methods
- Suggested solutions
"""

import sys
from pathlib import Path

def analyze_pdf(pdf_path):
    """Analyze PDF structure and content"""

    try:
        import pdfplumber
    except ImportError:
        print("❌ pdfplumber not installed")
        return

    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        print(f"❌ File not found: {pdf_path}")
        return

    print("\n" + "="*70)
    print("  PDF DIAGNOSTIC REPORT")
    print("="*70)

    print(f"\n📄 File: {pdf_path.name}")
    print(f"   Size: {pdf_path.stat().st_size / 1024:.1f} KB")
    print("-" * 70)

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            print(f"\n📊 PDF STRUCTURE")
            print(f"   Total pages: {len(pdf.pages)}")

            # Analyze each page
            for page_idx, page in enumerate(pdf.pages):
                print(f"\n   Page {page_idx + 1}:")
                print(f"   " + "-" * 60)

                # Check for tables
                tables = page.extract_tables()
                if tables:
                    print(f"      Tables found: {len(tables)}")
                    for table_idx, table in enumerate(tables):
                        print(f"      Table {table_idx + 1}: {len(table)} rows × {len(table[0]) if table else 0} cols")
                        # Show first few rows
                        print(f"         Header: {table[0][:3] if table else 'N/A'}...")
                else:
                    print(f"      Tables found: 0")

                # Check for text
                text = page.extract_text()
                if text and text.strip():
                    lines = text.split('\n')
                    print(f"      Text lines: {len(lines)}")
                    print(f"      Preview: {lines[0][:60]}...")
                else:
                    print(f"      Text lines: 0")

                # Check for rectangles/shapes
                rects = page.rects
                lines = page.lines
                curves = page.curves
                print(f"      Objects: {len(rects)} rectangles, {len(lines)} lines, {len(curves)} curves")

                # Show raw content types
                if hasattr(page, 'objects'):
                    obj_types = {}
                    for obj in page.objects.values():
                        for item in obj:
                            obj_type = item.get('object_type', 'unknown')
                            obj_types[obj_type] = obj_types.get(obj_type, 0) + 1
                    if obj_types:
                        print(f"      Content types: {obj_types}")

    except Exception as e:
        print(f"❌ Error analyzing PDF: {e}")
        import traceback
        traceback.print_exc()
        return

    # Recommendations
    print("\n" + "="*70)
    print("  RECOMMENDATIONS")
    print("="*70)

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            has_tables = False
            has_text = False

            for page in pdf.pages:
                if page.extract_tables():
                    has_tables = True
                if page.extract_text() and page.extract_text().strip():
                    has_text = True

            print()
            if has_tables:
                print("✓ Tables detected - Use 'Load Invoice File' in the app")
                print("  pdfplumber should extract them automatically")
            elif has_text:
                print("⚠️  No tables detected but text is present")
                print("   Options:")
                print("   1. Try OCR with custom template for text-based extraction")
                print("   2. Check if tables use borders/formatting pdfplumber can detect")
                print("   3. Use manual data entry or column mapping")
            else:
                print("❌ No tables or text detected")
                print("   This PDF appears to be image-based (scanned)")
                print("   Use OCR: python create_ocr_template.py")

    except Exception as e:
        print(f"Error: {e}")

    print("\n" + "="*70)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python debug_pdf.py <pdf_path>")
        print("Example: python debug_pdf.py Input/invoice.pdf")
        sys.exit(1)

    pdf_path = sys.argv[1]
    analyze_pdf(pdf_path)
