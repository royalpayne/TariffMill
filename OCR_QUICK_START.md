# OCR Quick Start Guide

## The Problem

You got "no valid table found in pdf" - this means your PDF is **scanned** (image-based), not digital text. OCR is needed to extract text from these images.

## Quick Test (2 minutes)

1. **Place a scanned PDF** in `DerivativeMill/Input/` folder
2. **Run the test:**
   ```bash
   cd DerivativeMill
   python test_ocr.py
   ```
3. **Review the output:**
   - If extraction works → Your template might be ready
   - If not → Create a custom template (see next section)

## Create Custom Template (5 minutes)

If the default template doesn't work, create one for your supplier:

```bash
cd DerivativeMill
python create_ocr_template.py
```

The script will:
1. Ask for your supplier name (e.g., "ACME Electronics")
2. Show you the extracted text
3. Guide you through setting regex patterns
4. Save the template automatically
5. Test the template on your invoice

## Template File Structure

Templates are stored as JSON in `DerivativeMill/ocr/templates/`:

```json
{
  "supplier_name": "ACME Electronics",
  "patterns": {
    "part_number_value": "([A-Z0-9\\-]{3,25})",
    "value_pattern": "\\$?\\s*(\\d+(?:[,\\.]\\d{1,3})*(?:\\.\\d{2})?)"
  },
  "field_positions": {}
}
```

**Key patterns:**
- `part_number_value`: Regex to find Part Numbers (e.g., ABC-123)
- `value_pattern`: Regex to find Prices (e.g., $49.99)

## Using in the App

Once your template is created:

1. Go to **"Invoice Mapping Profiles"** tab
2. Click **"Load Invoice File"**
3. Select your scanned PDF
4. OCR will automatically detect it's scanned and extract with your template
5. Review extracted columns
6. Map them to your fields

## Common Issues & Fixes

### Issue 1: "No Part Number/Value combinations found"

**Cause:** Your regex patterns don't match the invoice format

**Fix:**
1. Run `python test_ocr.py`
2. Look at the extracted text
3. Identify the actual format (layout, spacing, etc.)
4. Run `python create_ocr_template.py` again
5. Adjust patterns based on what you see

### Issue 2: "OCR found no text in image"

**Cause:** Image quality too low

**Fix:**
- Rescan invoice at 150+ DPI
- Check file size (should be 100KB+)
- Try a different invoice from the same supplier

### Issue 3: Extracted data is wrong

**Cause:** Patterns too broad or too narrow

**Fix:**
1. Edit the template JSON directly:
   ```bash
   nano DerivativeMill/ocr/templates/YourSupplier.json
   ```
2. Make patterns more specific
3. Test again with `python test_ocr.py`

## Regex Pattern Quick Reference

| Pattern | Matches | Example |
|---------|---------|---------|
| `[A-Z0-9\-]` | Letters, numbers, dashes | `ABC-123` |
| `{3,25}` | 3 to 25 characters | Part numbers |
| `\d+` | One or more digits | `123` |
| `\$` | Dollar sign | `$49.99` |
| `\.` | Decimal point | `49.99` |
| `,` | Comma (separator) | `1,234` |
| `\s+` | Whitespace | Spaces/tabs |
| `()` | Capture group | Returns the match |

## Template Examples

### Example 1: Standard Invoice
```
Part Number | Unit Price
ABC-123     | $49.99
XYZ-456     | $99.50
```

**Template:**
```json
{
  "part_number_value": "([A-Z0-9\\-]{3,20})",
  "value_pattern": "\\$\\s*(\\d+\\.\\d{2})"
}
```

### Example 2: Space-Separated
```
ABC-123  49.99
XYZ-456  99.50
```

**Template:**
```json
{
  "part_number_value": "([A-Z0-9\\-]+)\\s+",
  "value_pattern": "\\s+(\\d+\\.\\d{2})"
}
```

### Example 3: Item/Price Format
```
Item: ABC-123
Price: $49.99
Item: XYZ-456
Price: $99.50
```

**Template:**
```json
{
  "part_number_value": "Item:\\s*([A-Z0-9\\-]+)",
  "value_pattern": "Price:\\s*\\$(\\d+\\.\\d{2})"
}
```

## For Multiple Suppliers

Create separate templates for each:

```bash
python create_ocr_template.py  # Creates "Supplier A.json"
python create_ocr_template.py  # Creates "Supplier B.json"
python create_ocr_template.py  # Creates "Supplier C.json"
```

Each template is independent and stored in `ocr/templates/`

## For Testing

Use the test scripts:

```bash
# See what OCR extracts
python test_ocr.py

# Create or update a template
python create_ocr_template.py

# Manual testing (Python console)
python
>>> from ocr import preview_extraction
>>> preview = preview_extraction("Input/your_invoice.pdf")
>>> print(preview['text_preview'])
```

## Template Editing (Advanced)

Edit JSON directly:

```bash
nano DerivativeMill/ocr/templates/YourSupplier.json
```

**Remember:**
- Backslashes must be escaped: `\d` → `\\d`
- Valid JSON syntax required
- Test with `python test_ocr.py` after changes

## Still Having Issues?

See the full guide: `OCR_TEMPLATE_SETUP_GUIDE.md`

It includes:
- Detailed pattern reference
- Advanced troubleshooting
- How to test patterns
- Full API reference

## Summary

```
1. Place scanned PDF in Input/ folder
2. Run: python test_ocr.py
3. If it works → You're done!
4. If not → Run: python create_ocr_template.py
5. Test again → Adjust patterns if needed
6. Use in app → Load in "Invoice Mapping Profiles"
```

---

**Status:** OCR system ready to use
**Next:** Add your first scanned invoice and test it!
