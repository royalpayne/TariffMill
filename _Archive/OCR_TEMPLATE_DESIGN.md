# OCR Template Design for Invoice Field Recognition

## Executive Summary

Yes, there are multiple ways to build OCR templates that recognize Part Number and Value fields in scanned invoices. This document explores the best approaches for your use case.

---

## Option 1: pytesseract + Template Matching (Recommended for MVP)

### What It Is
Use Tesseract OCR with custom field matching templates to:
1. Extract all text from PDF
2. Find "Part Number" or similar headers
3. Extract values in predictable patterns
4. Return structured data

### Advantages
- ✅ No machine learning needed
- ✅ Works with scanned PDFs
- ✅ Can handle various invoice formats
- ✅ Fast (< 2 seconds per page)
- ✅ Customizable pattern matching

### Disadvantages
- ✗ Requires Tesseract system installation
- ✗ Less accurate than ML models
- ✗ Manual template creation per supplier
- ✗ Doesn't work with handwritten text

### Implementation Complexity
**Low to Medium** - Straightforward pattern matching

### Code Example
```python
def extract_fields_from_scanned_pdf(pdf_path, template):
    """
    Extract Part Number and Value from scanned invoice using template.

    Args:
        pdf_path: Path to scanned PDF
        template: OCRTemplate with field definitions

    Returns:
        DataFrame with extracted data
    """
    import pytesseract
    from pdf2image import convert_from_path
    import re

    images = convert_from_path(pdf_path)
    extracted_data = []

    for image in images:
        # Extract text from image
        text = pytesseract.image_to_string(image, lang='eng')

        # Find Part Number section
        part_num_section = template.find_section(text, 'part_number')
        if part_num_section:
            parts = template.extract_values(part_num_section, 'part_number')
            extracted_data.extend(parts)

    return pd.DataFrame(extracted_data)
```

---

## Option 2: OpenAI Vision API (Recommended for Accuracy)

### What It Is
Use ChatGPT's vision capabilities to intelligently extract data from invoice images with natural language instructions.

### Advantages
- ✅ Highest accuracy (95%+)
- ✅ Works with any layout
- ✅ Handles handwritten text
- ✅ No training needed
- ✅ One-shot learning (explain once, works for all)
- ✅ Can extract context-aware values

### Disadvantages
- ✗ Requires API key and cloud connection
- ✗ Cost per page ($0.01-0.03 per image)
- ✗ Network latency
- ✗ Privacy concerns (data sent to OpenAI)
- ✗ Rate limits

### Implementation Complexity
**Very Low** - Just send image and ask for JSON response

### Code Example
```python
def extract_fields_with_vision_api(pdf_path, api_key):
    """
    Extract Part Number and Value using OpenAI Vision API.
    """
    import requests
    from pdf2image import convert_from_path
    import base64

    images = convert_from_path(pdf_path)
    extracted_data = []

    for image in images:
        # Convert image to base64
        buffered = BytesIO()
        image.save(buffered, format="PNG")
        base64_image = base64.b64encode(buffered.getvalue()).decode()

        # Send to OpenAI Vision API
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": "gpt-4-vision-preview",
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                            },
                            {
                                "type": "text",
                                "text": "Extract all Part Numbers and their Values from this invoice. Return as JSON: {\"rows\": [{\"part_number\": \"...\", \"value\": \"...\"}, ...]}"
                            }
                        ]
                    }
                ]
            }
        )

        result = response.json()
        extracted_data.extend(result['choices'][0]['message']['content'])

    return pd.DataFrame(extracted_data)
```

---

## Option 3: ML-Based Field Detection (Advanced)

### What It Is
Train a custom ML model to recognize and extract Part Number and Value fields.

### Best Tools
1. **Document AI (Google Cloud)** - Pre-trained, fine-tunable
2. **AWS Textract** - Purpose-built for document extraction
3. **PyTorch + YOLO** - Custom computer vision model
4. **Hugging Face Models** - Pre-trained document understanding

### Advantages
- ✅ Highest accuracy with training
- ✅ Works with any layout
- ✅ Can handle complex documents
- ✅ Fast once trained

### Disadvantages
- ✗ High implementation complexity
- ✗ Requires training data (50-500 samples)
- ✗ Expensive cloud services
- ✗ Long development time (weeks)
- ✗ Ongoing maintenance

### Implementation Complexity
**Very High** - Requires ML expertise

---

## Option 4: Template-Based Pattern Recognition (Lightweight)

### What It Is
Define supplier-specific templates with regex patterns and field locations.

### Advantages
- ✅ No external dependencies
- ✅ Works offline
- ✅ Fast and reliable
- ✅ Customizable per supplier
- ✅ Version controllable

### Disadvantages
- ✗ Manual template creation
- ✗ Not flexible with layout changes
- ✗ Requires supplier-specific setup
- ✗ More error-prone with OCR

### Template Example
```python
class SupplierTemplate:
    def __init__(self, supplier_name):
        self.supplier = supplier_name
        self.patterns = {
            'part_number_header': r'(Part\s*Number|SKU|Product\s*ID)',
            'part_number_value': r'([A-Z0-9\-]{5,20})',
            'value_header': r'(Price|Unit\s*Price|Value|Amount)',
            'value_pattern': r'\$?\s*(\d+\.?\d*)',
        }
        self.field_positions = {
            'part_number': {'x': 50, 'y': 100, 'width': 200, 'height': 30},
            'value': {'x': 500, 'y': 100, 'width': 150, 'height': 30}
        }

    def extract(self, text, image=None):
        """Extract fields using patterns and positions."""
        results = []

        # Use regex patterns
        for pattern in self.patterns:
            matches = re.findall(self.patterns[pattern], text)
            # Process matches...

        return results
```

---

## Recommended Implementation Strategy

### Phase 3A: Scanned PDF Detection + pytesseract (MVP)

**Effort:** 1-2 weeks
**Cost:** Free (pytesseract)
**Accuracy:** 80-90%

```python
def handle_scanned_pdf(pdf_path):
    """
    Main entry point for scanned PDF handling.

    1. Detect if PDF is scanned
    2. Extract text with pytesseract
    3. Apply supplier template
    4. Return structured data
    """

    # Check if PDF is scanned
    if is_scanned_pdf(pdf_path):
        # Convert PDF to images
        images = pdf_to_images(pdf_path)

        # Extract text via OCR
        text = pytesseract.image_to_string(images[0])

        # Detect supplier
        supplier = detect_supplier(text)

        # Load template
        template = load_supplier_template(supplier)

        # Extract fields
        data = template.extract(text)

        return pd.DataFrame(data)
    else:
        # Fall back to pdfplumber table extraction
        return extract_pdf_table(pdf_path)
```

### Phase 3B: Smart Template Builder (Advanced)

Allow users to visually create templates:
- Upload sample invoice
- Click on Part Number field → Store pattern
- Click on Value field → Store pattern
- Save as supplier template
- Re-use for future invoices

### Phase 3C: AI-Powered Field Recognition (Future)

Use OpenAI Vision API for:
- One-shot learning (show one example)
- Extract from any invoice format
- Handle edge cases

---

## Comparison Matrix

| Approach | Accuracy | Speed | Cost | Complexity | Offline |
|----------|----------|-------|------|-----------|---------|
| **pytesseract** | 80-90% | Fast | Free | Low | ✅ Yes |
| **OpenAI Vision** | 95%+ | Slow | $0.01-0.03 | Very Low | ❌ No |
| **AWS Textract** | 95%+ | Medium | $0.015-0.04 | Low | ❌ No |
| **Custom ML** | 98%+ | Fast | High | Very High | ✅ Yes |
| **Templates** | 85-95% | Very Fast | Free | Medium | ✅ Yes |

---

## Recommended Path for Your Project

### For MVP (Quick Implementation)
**Use: pytesseract + Simple Pattern Matching**

```
Effort: 1-2 weeks
Cost: Free
Result: Works with most standard invoices
```

**Steps:**
1. Install pytesseract + Tesseract
2. Add scanned PDF detection
3. Create 2-3 supplier templates
4. Test with scanned samples

### For Phase 3B (Enhanced UX)
**Add: Template Builder UI**

```
Effort: 2-3 weeks
Cost: Free
Result: Users can create templates themselves
```

**Features:**
- Visual template editor
- Pattern matching examples
- Test on sample invoices
- Save/load templates

### For Phase 3C (Ultimate Solution)
**Use: OpenAI Vision API (Optional)**

```
Effort: 1 week
Cost: $0.01-0.03 per page
Result: Works with any invoice format
```

**Benefits:**
- Single solution for all formats
- No templates needed
- Highest accuracy
- Easiest for users

---

## Implementation Priority

### 🥇 Priority 1: Scanned PDF Detection
```python
def is_scanned_pdf(pdf_path):
    """Determine if PDF is scanned or digital."""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Digital PDFs have searchable text
            if page.extract_text():
                return False
    return True  # No text found = scanned
```

### 🥈 Priority 2: pytesseract Integration
```python
def extract_from_scanned_invoice(pdf_path):
    """Extract text from scanned invoice."""
    images = pdf_to_images(pdf_path)
    text = pytesseract.image_to_string(images[0])
    return text
```

### 🥉 Priority 3: Pattern-Based Field Extraction
```python
def extract_fields_from_text(text, supplier_name):
    """Extract Part Number and Value from OCR text."""
    template = load_template(supplier_name)
    return template.extract(text)
```

---

## System Requirements

### For pytesseract Approach
```bash
# System-level dependency
apt-get install tesseract-ocr  # Linux
brew install tesseract         # macOS
# Or: Download from https://github.com/UB-Mannheim/tesseract/wiki

# Python packages
pip install pytesseract pdf2image pillow
```

### For OpenAI Vision Approach
```bash
pip install openai requests
# Requires: OpenAI API key
```

### For AWS Textract Approach
```bash
pip install boto3
# Requires: AWS credentials and account
```

---

## Code Organization

### Suggested Module Structure
```
DerivativeMill/
├── ocr/
│   ├── __init__.py
│   ├── scanned_pdf.py          # Detection logic
│   ├── pytesseract_ocr.py      # Tesseract integration
│   ├── template_matching.py    # Pattern-based extraction
│   ├── field_detector.py       # Field recognition
│   ├── supplier_templates/
│   │   ├── default.py
│   │   ├── supplier_a.py
│   │   └── supplier_b.py
│   └── templates/
│       ├── supplier_a.json
│       └── supplier_b.json
└── derivativemill.py
```

---

## Integration with Existing Code

### Modified `load_csv_for_shipment_mapping()`
```python
def load_csv_for_shipment_mapping(self):
    path, _ = QFileDialog.getOpenFileName(...)

    if path.lower().endswith('.pdf'):
        with pdfplumber.open(path) as pdf:
            # Check if digital or scanned
            if is_scanned_pdf(pdf):
                # New OCR path
                df = extract_from_scanned_invoice(path)
            else:
                # Existing pdfplumber path
                df = self.extract_pdf_table(path)
```

### New `is_scanned_pdf()` Function
```python
def is_scanned_pdf(pdf):
    """Check if PDF contains extractable text."""
    for page in pdf.pages:
        if page.extract_text(layout=False).strip():
            return False  # Digital PDF
    return True  # Scanned PDF
```

---

## Testing Strategy

### Unit Tests
```python
def test_scanned_pdf_detection():
    # Test with digital PDF → Should return False
    # Test with scanned PDF → Should return True

def test_ocr_field_extraction():
    # Test with scanned invoice sample
    # Verify Part Number extracted
    # Verify Value extracted
    # Check accuracy (>80%)

def test_supplier_template():
    # Load supplier template
    # Extract fields
    # Compare with expected results
```

### Integration Tests
```python
def test_full_scanned_invoice_workflow():
    # Load scanned PDF
    # Detect it's scanned
    # Extract fields via OCR
    # Map to Part Number and Value
    # Verify mappings are correct
    # Save as profile
```

---

## Known Limitations & Workarounds

### Limitation 1: Low OCR Accuracy with Complex Layouts
**Workaround:** User can fall back to CSV/Excel or manual entry

### Limitation 2: Handwritten Text Not Recognized
**Workaround:** Use OpenAI Vision API for handwritten support

### Limitation 3: Tables Not Extracted from Scanned PDFs
**Workaround:** Tesseract extracts all text; patterns find structure

### Limitation 4: Supplier-Specific Templates Needed
**Workaround:** Build template for each supplier once, then re-use

---

## Cost Analysis

### Free Option (pytesseract)
```
Cost: $0
Effort: 20-30 hours
Accuracy: 80-90%
Best For: Standard invoices from known suppliers
```

### Low-Cost Option (OpenAI Vision)
```
Cost: $5-20/month (typical usage)
Effort: 4-8 hours
Accuracy: 95%+
Best For: Any invoice format
```

### Enterprise Option (AWS Textract)
```
Cost: $0.015 per page (bulk discount available)
Effort: 8-12 hours
Accuracy: 95%+
Best For: High-volume, critical documents
```

---

## Recommendation Summary

### **For Your Project (MVP Phase 3)**

1. **Start with:** pytesseract + scanned PDF detection
   - Easy to implement
   - Works offline
   - Free
   - Handles most invoices

2. **Create templates** for your top 3-5 suppliers
   - Store regex patterns
   - Store field positions
   - JSON-based for easy maintenance

3. **Add UI** to warn users about scanned PDFs
   - Show extracted text
   - Allow manual correction
   - Save corrections to template

4. **Future:** Consider OpenAI Vision API if:
   - Many different suppliers
   - Accuracy critical
   - Want one-click extraction

---

## Next Steps

1. **Decide:** Which approach suits your needs?
2. **Prototype:** Test pytesseract with sample scanned invoices
3. **Design:** Create template format for your suppliers
4. **Implement:** Add OCR to Invoice Mapping tab
5. **Deploy:** Gather feedback, iterate

---

## Questions to Ask Before Implementation

1. **Do you have scanned invoices?** (If no, pdfplumber is enough)
2. **How many suppliers?** (More = template builder more valuable)
3. **Accuracy requirement?** (> 95% = use Vision API)
4. **Budget for APIs?** (Yes = use OpenAI; No = use pytesseract)
5. **User comfort with tech?** (Low = vision API easier)

---

## References

- **pytesseract:** https://github.com/madmaze/pytesseract
- **pdf2image:** https://github.com/Belval/pdf2image
- **OpenAI Vision:** https://platform.openai.com/docs/guides/vision
- **AWS Textract:** https://aws.amazon.com/textract/
- **Tesseract OCR:** https://github.com/UB-Mannheim/tesseract/wiki

---

**Conclusion:** Yes, OCR templates are practical and buildable. Start with pytesseract for MVP, then enhance based on real-world usage and feedback.
