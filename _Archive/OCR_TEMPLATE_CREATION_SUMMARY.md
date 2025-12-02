# OCR Template Creation - Complete Summary

## Problem You Encountered

**Error:** "no valid table found in pdf"

**Reason:** Your PDF is scanned (image-based), not a digital text PDF. pdfplumber can't extract tables from images - you need OCR to extract text first.

**Good News:** The OCR system is fully integrated and ready to use!

---

## How OCR Works in DerivativeMill

```
Scanned Invoice PDF
    ↓
[OCR Integration detects: is this scanned?]
    ↓ YES (Scanned) → Use OCR
    ↓ NO (Digital)  → Use pdfplumber
    ↓
[Convert PDF to images]
    ↓
[Run Tesseract OCR to extract text]
    ↓
[Apply supplier template patterns]
    ↓
[Extract Part Numbers & Values]
    ↓
[Return structured DataFrame]
```

---

## Tools You Now Have

### 1. **test_ocr.py** - Test OCR Extraction

**What it does:** Shows you what OCR extracts from a scanned PDF

**How to use:**
```bash
cd DerivativeMill
python test_ocr.py
```

**Output:**
- Whether the PDF is scanned
- What text OCR extracted
- How many rows were extracted
- Saves result to CSV for review

### 2. **create_ocr_template.py** - Create Custom Templates

**What it does:** Interactive wizard to create supplier-specific templates

**How to use:**
```bash
cd DerivativeMill
python create_ocr_template.py
```

**Workflow:**
1. You select a scanned PDF
2. Script shows you the extracted text
3. You enter your supplier name
4. Script asks you to confirm regex patterns
5. Template is created and tested automatically
6. Template is saved to `ocr/templates/YourSupplier.json`

**Result:** A template file that recognizes fields in your supplier's invoices

### 3. **OCR_QUICK_START.md** - 5-Minute Guide

Quick reference with:
- How to test OCR (2 minutes)
- How to create templates (5 minutes)
- Common issues and fixes
- Template examples
- Regex pattern reference

### 4. **OCR_TEMPLATE_SETUP_GUIDE.md** - Comprehensive Guide

Detailed documentation with:
- Understanding template structure
- Manual template creation steps
- Pattern matching examples
- Regex reference
- Advanced troubleshooting
- API reference

---

## Quick Start Workflow

### Step 1: Test OCR (2 minutes)

```bash
cd DerivativeMill
python test_ocr.py
```

- Place your scanned PDF in `Input/` folder
- Edit PDF_PATH in test_ocr.py to match your file
- Run the script
- Check if extraction works

### Step 2: Create Template (5 minutes)

If default template doesn't work for your supplier:

```bash
cd DerivativeMill
python create_ocr_template.py
```

- Follows an interactive flow
- Asks for supplier name
- Shows extracted text
- Guides you through pattern configuration
- Tests the template automatically
- Saves to `ocr/templates/`

### Step 3: Use in App

1. Go to **"Invoice Mapping Profiles"** tab
2. Click **"Load Invoice File"**
3. Select your scanned PDF
4. OCR automatically detects it's scanned
5. Extraction completes with your template
6. Review extracted columns
7. Map to your fields

---

## Understanding Templates

### What is a Template?

A template tells OCR how to find Part Numbers and Values in YOUR supplier's invoice format.

### Template File Example

File: `DerivativeMill/ocr/templates/ACME Electronics.json`

```json
{
  "supplier_name": "ACME Electronics",
  "patterns": {
    "part_number_header": "(part\\s*#|part\\s*number|sku)",
    "part_number_value": "([A-Z0-9\\-]{3,20})",
    "value_header": "(unit\\s*price|price|amount)",
    "value_pattern": "\\$?\\s*(\\d+(?:[,\\.]\\d{1,3})*(?:\\.\\d{2})?)"
  },
  "field_positions": {}
}
```

### Key Parts

- **supplier_name**: Your supplier's name (identifies the template)
- **part_number_value**: Regex pattern to find Part Numbers
- **value_pattern**: Regex pattern to find Prices

### Simple vs. Complex Patterns

**Simple (works for most):**
```json
"part_number_value": "([A-Z0-9\\-]{3,20})",
"value_pattern": "\\$?\\s*(\\d+\\.\\d{2})"
```

**Complex (for unusual formats):**
```json
"part_number_value": "Item:\\s*([A-Z0-9\\-]+)",
"value_pattern": "Price:\\s*\\$(\\d+\\.\\d{2})"
```

---

## Common Patterns

### Pattern 1: Standard Invoice

```
Part Number | Unit Price
ABC-123     | $49.99
XYZ-456     | $99.50
```

**Template:**
```json
"part_number_value": "([A-Z0-9\\-]{3,20})",
"value_pattern": "\\$\\s*(\\d+\\.\\d{2})"
```

### Pattern 2: Space-Separated

```
ABC-123 49.99
XYZ-456 99.50
```

**Template:**
```json
"part_number_value": "([A-Z0-9\\-]+)\\s+",
"value_pattern": "\\s+(\\d+\\.\\d{2})"
```

### Pattern 3: Item/Price Labels

```
Item: ABC-123
Price: $49.99
Item: XYZ-456
Price: $99.50
```

**Template:**
```json
"part_number_value": "Item:\\s*([A-Z0-9\\-]+)",
"value_pattern": "Price:\\s*\\$(\\d+\\.\\d{2})"
```

---

## Template Creation Process

### Process Flow

```
1. Place scanned PDF in Input/
    ↓
2. Run: python test_ocr.py
    ↓
3. If it works → Done!
    ↓ If not...
4. Run: python create_ocr_template.py
    ↓
5. Enter supplier name
    ↓
6. Review extracted text
    ↓
7. Confirm/adjust patterns
    ↓
8. Template saved to ocr/templates/
    ↓
9. Test again with python test_ocr.py
    ↓
10. If still not perfect → Edit JSON directly
```

### Editing Templates Manually

If you need to fine-tune a template:

```bash
# List existing templates
ls -la DerivativeMill/ocr/templates/

# Edit a template
nano DerivativeMill/ocr/templates/YourSupplier.json

# Test after editing
python test_ocr.py
```

---

## Multiple Suppliers

Create separate templates for each supplier:

```bash
python create_ocr_template.py  # Creates template for Supplier A
python create_ocr_template.py  # Creates template for Supplier B
python create_ocr_template.py  # Creates template for Supplier C
```

Each template is stored independently in `ocr/templates/`

---

## Troubleshooting

### Problem 1: "No Part Number/Value combinations found"

**Cause:** Your patterns don't match the invoice text

**Solution:**
1. Look at the text preview from `test_ocr.py`
2. See what the actual format is
3. Adjust patterns in the template
4. Example: If text shows "SKU: ABC123" but pattern looks for dashes, update it

### Problem 2: "OCR found no text in image"

**Cause:** Image quality too low

**Solution:**
- Rescan at 150+ DPI
- Check file is actually a PDF (not corrupted)
- Try another invoice from the same supplier

### Problem 3: Wrong data extracted

**Cause:** Pattern too broad or too narrow

**Solution:**
1. Review extracted text carefully
2. Identify exact format
3. Make pattern more specific
4. Test again

---

## Integration in the App

When you load a scanned PDF:

1. **Detection:** `is_scanned_pdf()` checks if PDF is scanned
2. **Routing:**
   - Scanned → Uses OCR with your template
   - Digital → Uses pdfplumber
3. **Extraction:** Your template's patterns are applied
4. **Display:** Extracted columns appear as draggable labels
5. **Mapping:** You map to your fields

---

## System Architecture

```
DerivativeMill/
├── derivativemill.py          # Main app (now with OCR detection)
├── ocr/                        # OCR module
│   ├── __init__.py
│   ├── scanned_pdf.py         # PDF detection & conversion
│   ├── field_detector.py      # Template & pattern matching
│   ├── ocr_extract.py         # Main extraction pipeline
│   └── templates/             # Supplier templates
│       ├── ACME Electronics.json
│       ├── Widget Corp.json
│       └── ...
├── test_ocr.py                # Test tool
├── create_ocr_template.py     # Template creator tool
└── ...
```

---

## Key Points

✅ **OCR is fully integrated** into the app
✅ **Automatic detection** of scanned vs. digital PDFs
✅ **Template system** for supplier-specific patterns
✅ **Tools provided** for testing and creating templates
✅ **Fallback to pdfplumber** for digital PDFs
✅ **No impact** on existing CSV/Excel workflows

---

## Next Steps

### For You:
1. Get a sample scanned invoice from your supplier
2. Run `python test_ocr.py` to see what happens
3. Create a template with `python create_ocr_template.py`
4. Test in the app with the Invoice Mapping Profiles tab

### For Multiple Suppliers:
1. Create a template for each unique supplier format
2. Store them in `ocr/templates/`
3. Use in app - templates are auto-detected by supplier name

### For Production:
1. Test with real invoices
2. Fine-tune templates as needed
3. Merge feature branch to master when ready
4. Deploy to production

---

## Files Created

- `DerivativeMill/test_ocr.py` - Testing tool (215 lines)
- `DerivativeMill/create_ocr_template.py` - Template creator (280 lines)
- `OCR_QUICK_START.md` - Quick reference guide
- `OCR_TEMPLATE_SETUP_GUIDE.md` - Comprehensive guide
- `OCR_TEMPLATE_CREATION_SUMMARY.md` - This file

---

## Support Resources

| File | Purpose | Length |
|------|---------|--------|
| `OCR_QUICK_START.md` | 5-minute quick start | 250 lines |
| `OCR_TEMPLATE_SETUP_GUIDE.md` | Comprehensive guide | 600+ lines |
| `test_ocr.py` | Testing tool | 215 lines |
| `create_ocr_template.py` | Interactive template creator | 280 lines |
| `OCR_IMPLEMENTATION_SUMMARY.md` | API reference | 465 lines |

---

## Commit History

- `90689f0` - Add OCR template creation tools and comprehensive guides
- `96e0e43` - Integrate OCR scanned PDF detection into invoice loading
- `5e6d304` - OCR module implementation complete
- `494d5c9` - Implement pytesseract OCR system

---

**Status:** OCR system complete and ready for use

**Next Action:** Test with your first scanned invoice using `python test_ocr.py`
