# OCR Template Setup Guide for DerivativeMill

## Understanding the Problem

When you get "no valid table found in pdf", it means:
- **Your PDF is scanned** (image-based, not text-based)
- pdfplumber can't extract tables from images
- **OCR must be used** to extract text from the scanned image first

## Quick Start: Test OCR is Working

### Step 1: Create a Test Script

Create a file called `test_ocr.py` in the DerivativeMill folder:

```python
#!/usr/bin/env python3
"""Test OCR extraction and preview"""

from pathlib import Path
from ocr import preview_extraction, extract_from_scanned_invoice

# Replace with your actual scanned PDF path
pdf_path = Path("Input/your_scanned_invoice.pdf")

if not pdf_path.exists():
    print(f"❌ File not found: {pdf_path}")
    print("Please add a scanned invoice to the Input folder")
    exit(1)

print(f"\n📄 Testing OCR on: {pdf_path.name}")
print("=" * 60)

# Step 1: Preview the extracted text
print("\n1️⃣  PREVIEWING OCR TEXT EXTRACTION")
print("-" * 60)
preview = preview_extraction(str(pdf_path), max_lines=30)
print(f"Is Scanned: {preview['is_scanned']}")
print(f"Total lines: {preview['line_count']}")
print(f"Total characters: {preview['char_count']}")
print(f"\nExtracted Text Preview:")
print(preview['text_preview'])

# Step 2: Try full extraction
print("\n2️⃣  ATTEMPTING FULL EXTRACTION")
print("-" * 60)
try:
    df, metadata = extract_from_scanned_invoice(str(pdf_path))
    print(f"✅ Extraction successful!")
    print(f"   Rows extracted: {len(df)}")
    print(f"   Columns: {metadata['columns']}")
    print(f"\nFirst few rows:")
    print(df.head(10))
except Exception as e:
    print(f"❌ Extraction failed: {e}")
    print("\nThis is expected if OCR patterns don't match your invoice format.")
    print("Follow the template customization steps below.")
```

### Step 2: Run the Test

```bash
cd DerivativeMill
python test_ocr.py
```

## Creating OCR Templates

### Understanding Template Structure

Each supplier invoice has a unique layout. OCR templates contain **regex patterns** to find Part Numbers and Values.

Default pattern locations:
- **Part Numbers**: `SKU-123`, `A-456`, `PROD789` (3-25 alphanumeric characters)
- **Values**: `$49.99`, `100.50`, `1,234.56` (prices with optional $ and commas)

### Method 1: Manual Template Creation (Recommended for First-Time)

#### Step 1: Analyze Your Invoice

```python
from ocr import preview_extraction

pdf_path = "Input/your_supplier_invoice.pdf"
preview = preview_extraction(pdf_path, max_lines=50)

# Copy the text_preview output and analyze it
# Look for patterns like:
# - Where part numbers appear (format, spacing)
# - Where prices appear (format, currency symbol)
```

#### Step 2: Create Supplier Template

Create a Python script `create_template.py`:

```python
#!/usr/bin/env python3
"""Create a custom OCR template for your supplier"""

from ocr import SupplierTemplate, get_template_manager

# Step 1: Define your supplier name
SUPPLIER_NAME = "ACME Electronics"  # Change this to your supplier

# Step 2: Create a new template
template = SupplierTemplate(SUPPLIER_NAME)

# Step 3: Customize patterns based on your invoice format
# These patterns help OCR find the Part Number and Value columns

# Example 1: If your invoice has a header like "Part # | Unit Price"
template.patterns['part_number_header'] = r'(part\s*#|part\s*number|sku|product\s*id)'
template.patterns['part_number_value'] = r'([A-Z0-9\-]{3,20})'
template.patterns['value_header'] = r'(unit\s*price|price|amount|cost)'
template.patterns['value_pattern'] = r'\$?\s*(\d+(?:[,\.]\d{3})*(?:\.\d{2})?)'

# Example 2: If your invoice uses different patterns, modify as needed
# (See "Pattern Reference" section below)

# Step 4: Save the template
manager = get_template_manager()
manager.save_template(template)

print(f"✅ Template saved for {SUPPLIER_NAME}")
print(f"   Location: DerivativeMill/ocr/templates/{SUPPLIER_NAME}.json")
```

Run it:
```bash
python create_template.py
```

#### Step 3: Test Your Template

```python
from ocr import extract_from_scanned_invoice

pdf_path = "Input/your_supplier_invoice.pdf"
supplier_name = "ACME Electronics"  # Must match your template name

try:
    df, metadata = extract_from_scanned_invoice(pdf_path, supplier_name=supplier_name)
    print(f"✅ Success! Extracted {len(df)} rows")
    print(df.head())
except Exception as e:
    print(f"❌ Failed: {e}")
    print("Your patterns may not match the invoice format.")
    print("Try adjusting the regex patterns in create_template.py")
```

## Pattern Reference

### Common Invoice Layouts

#### Layout 1: Table with Headers
```
Part Number | Unit Price | Qty
ABC-123     | $49.99     | 10
XYZ-456     | $99.50     | 5
```

**Template Settings:**
```python
template.patterns['part_number_value'] = r'([A-Z0-9\-]{3,20})'
template.patterns['value_pattern'] = r'\$\s*(\d+\.\d{2})'
```

#### Layout 2: Space-Separated Values
```
ABC-123  49.99
XYZ-456  99.50
```

**Template Settings:**
```python
template.patterns['part_number_value'] = r'([A-Z0-9\-]{3,20})\s+(\d+\.\d{2})'
template.patterns['value_pattern'] = r'(\d+\.\d{2})'
```

#### Layout 3: Line Item Format
```
Item: ABC-123
Price: $49.99
---
Item: XYZ-456
Price: $99.50
```

**Template Settings:**
```python
template.patterns['part_number_header'] = r'item:?'
template.patterns['part_number_value'] = r'([A-Z0-9\-]{3,20})'
template.patterns['value_header'] = r'price:?'
template.patterns['value_pattern'] = r'\$(\d+\.\d{2})'
```

### Regex Pattern Symbols

| Symbol | Meaning | Example |
|--------|---------|---------|
| `\d` | Any digit (0-9) | `\d{2}` = "49" |
| `+` | One or more | `\d+` = "123" or "1" |
| `*` | Zero or more | `[A-Z]*` = "ABC" or "" |
| `{n}` | Exactly n times | `\d{2}` = "49" |
| `{n,m}` | Between n and m | `\d{1,2}` = "5" or "50" |
| `[A-Z]` | Any letter A-Z | `[A-Z0-9]` = "A" or "5" |
| `\s` | Whitespace | `\s+` = spaces/tabs |
| `\$` | Dollar sign | `\$49` |
| `\.` | Decimal point | `49\.99` |
| `,` | Comma | `1,234` |
| `()` | Capture group | `(\d+)` = captures the number |

## Template File Structure

Templates are stored as JSON files in `DerivativeMill/ocr/templates/`:

**Example: `ACME Electronics.json`**
```json
{
  "supplier_name": "ACME Electronics",
  "patterns": {
    "part_number_header": "(part\\s*#|part\\s*number|sku)",
    "part_number_value": "([A-Z0-9\\-]{3,20})",
    "value_header": "(unit\\s*price|price|amount)",
    "value_pattern": "\\$?\\s*(\\d+(?:[,\\.]\\d{3})*(?:\\.\\d{2})?)"
  },
  "field_positions": {}
}
```

**Note:** Backslashes are escaped in JSON (`\\` instead of `\`)

## Advanced: Modify Template Directly

Edit `DerivativeMill/ocr/templates/ACME Electronics.json`:

```bash
# View existing templates
ls -la DerivativeMill/ocr/templates/

# Edit a template (open in your text editor)
cat DerivativeMill/ocr/templates/YourSupplier.json
```

## Troubleshooting

### Problem: "No Part Number/Value combinations found"

**Cause:** Your regex patterns don't match the actual invoice text

**Solution:**
1. Run `test_ocr.py` and save the text preview to a file
2. Manually inspect where part numbers and values appear
3. Adjust regex patterns to match the actual format

Example:
```python
# If OCR extracted text looks like:
# "SKU: ABC123   COST: 49.99"
# But your patterns look for dashes and dollar signs

# Change this:
template.patterns['part_number_value'] = r'([A-Z0-9\-]{3,20})'  # Expects dashes

# To this:
template.patterns['part_number_value'] = r'SKU:\s*([A-Z0-9]+)'   # Matches "SKU: ABC123"
```

### Problem: "OCR found no text in image"

**Cause:** Image quality is too low or PDF is blank

**Solutions:**
1. Check file size - scanned PDFs should be 100KB+
2. Try rescanning at 150+ DPI
3. Verify the PDF opens in Adobe Reader

### Problem: Wrong columns extracted

**Cause:** Multiple tables detected or pattern is too broad

**Solutions:**
1. Make patterns more specific
2. Add supplier-specific header detection
3. Manually verify extracted text with `preview_extraction()`

## Using Templates in the Application

Once your template is created and saved:

1. **Automatic Detection** (if supplier name matches):
   - When you load a scanned PDF, the app will auto-detect the supplier
   - If no match, uses default template

2. **Manual Selection**:
   - In future UI updates, you'll be able to select the supplier template
   - Currently uses "default" template

## Testing Workflow

```
1. Add scanned PDF to Input/ folder
2. Run test_ocr.py to see extracted text
3. Create template for your supplier
4. Run test again to verify accuracy
5. Load in app using Invoice Mapping Profiles tab
6. Verify extracted columns match your data
```

## Next Steps

### For Testing:
1. Create a test script with your sample invoice
2. Try extraction with default template
3. Create supplier-specific template if needed

### For Production:
1. Create templates for each supplier
2. Store templates in `ocr/templates/`
3. Share template files with team
4. Update when supplier changes format

## Resources

- **OCR Module API**: See `OCR_IMPLEMENTATION_SUMMARY.md`
- **Pattern Testing**: https://regex101.com (test patterns here)
- **Tesseract OCR**: Installed in venv, handles text recognition

## Questions?

1. **Pattern not matching?** → Use `preview_extraction()` to see exact text
2. **OCR accuracy low?** → Check image quality, rescan at 150+ DPI
3. **Multiple suppliers?** → Create separate templates for each
4. **Template format wrong?** → Check JSON syntax with `python -m json.tool`

---

**Status:** OCR system ready for template configuration
**Next:** Create templates for your specific suppliers
