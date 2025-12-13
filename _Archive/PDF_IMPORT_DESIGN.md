# PDF Invoice Import Feature Design
## DerivativeMill v1.09+

---

## Executive Summary

Add PDF invoice import capability to the **Invoice Mapping Profiles** tab with seamless integration into the existing drag-and-drop mapping workflow. The feature will extract tabular data from PDFs and present it in the same format as CSV/Excel files.

---

## Architecture Overview

### 1. **Library Selection: PyPDF2 or pdfplumber**

| Library | Pros | Cons | Best For |
|---------|------|------|----------|
| **pdfplumber** | ✓ Better table extraction ✓ Accurate layout preservation ✓ Easy to implement | Slightly larger dependency | **RECOMMENDED** |
| **PyPDF2** | ✓ Lightweight ✓ Already common | ✗ Poor table extraction | Basic text only |

**Recommendation:** Use **pdfplumber** for reliable table extraction

### Install:
```bash
pip install pdfplumber
```

---

## Implementation Strategy

### Phase 1: Minimal Implementation (MVP)
**Scope:** Add PDF support to existing Invoice Mapping tab with minimal code changes

#### Option 1: Single File (Recommended for now)
```
load_csv_for_shipment_mapping()
├── Check file extension (.pdf, .csv, .xlsx)
├── If PDF:
│   └── extract_table_from_pdf()
│       └── Returns DataFrame
└── If CSV/Excel:
    └── Existing logic
```

#### Option 2: Separate Method (Better long-term)
```
load_csv_for_shipment_mapping()
├── File dialog
└── Route to appropriate handler:
    ├── _load_spreadsheet_for_mapping()
    ├── _load_pdf_for_mapping()
    └── _extract_dataframe_from_source()
```

---

## Detailed Implementation (Option 1 - Recommended)

### Step 1: Update File Dialog Filter
**Location:** Line 2682 in `load_csv_for_shipment_mapping()`

```python
# BEFORE:
path, _ = QFileDialog.getOpenFileName(self, "Select CSV/Excel", str(INPUT_DIR), "CSV/Excel Files (*.csv *.xlsx)")

# AFTER:
path, _ = QFileDialog.getOpenFileName(
    self,
    "Select Invoice File",
    str(INPUT_DIR),
    "All Supported (*.csv *.xlsx *.pdf);;CSV Files (*.csv);;Excel Files (*.xlsx);;PDF Files (*.pdf)"
)
```

### Step 2: Add PDF Extraction Function
**Location:** New function after `load_csv_for_shipment_mapping()`

```python
def extract_pdf_table(self, pdf_path):
    """
    Extract tabular data from PDF invoices

    Args:
        pdf_path: Path to PDF file

    Returns:
        DataFrame with extracted data, or None if extraction fails

    Raises:
        Exception: If PDF cannot be processed
    """
    try:
        import pdfplumber

        with pdfplumber.open(pdf_path) as pdf:
            # Strategy: Try each page until we find a table
            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                if tables and len(tables) > 0:
                    # Use the first (largest) table
                    table = max(tables, key=len)

                    if len(table) > 0:
                        # Convert to DataFrame
                        headers = table[0]
                        data = table[1:]

                        # Filter out empty rows
                        data = [row for row in data if any(cell for cell in row)]

                        if not data:
                            continue

                        df = pd.DataFrame(data, columns=headers)

                        logger.info(f"PDF table extracted from page {page_idx + 1}: {df.shape}")
                        return df

            raise ValueError("No valid table found in PDF")

    except ImportError:
        raise Exception("pdfplumber not installed. Run: pip install pdfplumber")
    except Exception as e:
        raise Exception(f"PDF extraction failed: {str(e)}")
```

### Step 3: Update Load Function
**Location:** Update `load_csv_for_shipment_mapping()` Line 2681-2702

```python
def load_csv_for_shipment_mapping(self):
    path, _ = QFileDialog.getOpenFileName(
        self,
        "Select Invoice File",
        str(INPUT_DIR),
        "All Supported (*.csv *.xlsx *.pdf);;CSV Files (*.csv);;Excel Files (*.xlsx);;PDF Files (*.pdf)"
    )
    if not path:
        return

    try:
        # Determine file type and extract data
        file_ext = Path(path).suffix.lower()

        if file_ext == '.pdf':
            df = self.extract_pdf_table(path)
        elif file_ext == '.xlsx':
            df = pd.read_excel(path, nrows=0, dtype=str)
        else:  # .csv
            df = pd.read_csv(path, nrows=0, dtype=str)

        cols = list(df.columns)

        # Clear existing labels
        for label in self.shipment_drag_labels:
            label.setParent(None)
        self.shipment_drag_labels = []

        # Add new labels from extracted columns
        left_layout = self.shipment_widget.layout().itemAt(0).widget().layout()
        for col in cols:
            lbl = DraggableLabel(col)
            left_layout.insertWidget(left_layout.count()-1, lbl)
            self.shipment_drag_labels.append(lbl)

        file_type = "PDF" if file_ext == '.pdf' else ("Excel" if file_ext == '.xlsx' else "CSV")
        logger.info(f"{file_type} file loaded for mapping: {Path(path).name}")
        self.status.setText(f"{file_type} file loaded: {Path(path).name}")

    except Exception as e:
        QMessageBox.critical(self, "Error", f"Cannot read file:\n{e}")
        logger.error(f"File loading failed: {str(e)}")
```

---

## Phase 2: Advanced Features (Future)

### Feature 2A: Multi-table Selection
If PDF has multiple tables, let user choose which one to use:

```python
def extract_pdf_table_advanced(self, pdf_path):
    """Extract with table selection UI"""
    import pdfplumber

    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables or []):
                if table and len(table) > 1:
                    all_tables.append({
                        'data': table,
                        'page': page_idx + 1,
                        'table': table_idx + 1,
                        'rows': len(table)
                    })

    if not all_tables:
        raise ValueError("No tables found in PDF")

    if len(all_tables) == 1:
        return pd.DataFrame(all_tables[0]['data'][1:], columns=all_tables[0]['data'][0])

    # If multiple tables, show selection dialog
    return self._show_table_selection_dialog(all_tables)
```

### Feature 2B: Column Name Cleanup
Auto-clean extracted column names (remove extra whitespace, normalize):

```python
def clean_column_names(self, df):
    """Clean extracted column names"""
    df.columns = [col.strip().lower().replace('\n', ' ') for col in df.columns]
    return df
```

### Feature 2C: Data Type Detection
Auto-detect and suggest column mappings:

```python
def auto_suggest_mappings(self, df):
    """Suggest column mappings based on content analysis"""
    suggestions = {}

    for col in df.columns:
        col_lower = col.lower()

        if any(term in col_lower for term in ['part', 'number', 'sku', 'product']):
            suggestions['part_number'] = col
        elif any(term in col_lower for term in ['value', 'price', 'amount', 'usd', 'cost']):
            suggestions['value_usd'] = col

    return suggestions
```

---

## Integration Points

### User Interface Changes

#### 1. Button Label Update
```python
# Line 2632 - Change button text to reflect PDF support
btn_load_csv = QPushButton("Load Invoice File")  # Was "Load CSV to Map"
```

#### 2. Update User Guide
In the User Guide tab, update step 3:

```html
<h3>Step 3: Create Invoice Mapping Profiles</h3>
<div class="workflow">
    <b>Location:</b> <span class="button-text">Invoice Mapping Profiles</span> tab<br>
    <div class="workflow-step"><span class="button-text">Load Invoice File</span> - Select CSV, Excel, or PDF invoice
        <ul style="margin-top: 8px;">
            <li><b>CSV/Excel:</b> Auto-detects column headers</li>
            <li><b>PDF:</b> Extracts tables and uses first valid table</li>
        </ul>
    </div>
    ...
</div>
```

---

## Error Handling Strategy

### Error Scenarios & Solutions

| Scenario | Solution |
|----------|----------|
| PDF has no tables | Show error: "No tables found in PDF. Try a different file." |
| PDF extraction fails | Show error: "Cannot extract PDF data. Ensure it's a standard invoice format." |
| pdfplumber not installed | Show error with pip install command |
| Multiple tables in PDF | Use largest table (or offer selection in Phase 2) |
| Corrupted PDF | Catch exception: "PDF appears corrupted. Try opening in Adobe Reader first." |
| Scanned PDF (image) | Show info: "PDF appears to be scanned. Consider using a digital invoice instead." |

### Implementation:
```python
def extract_pdf_table(self, pdf_path):
    try:
        import pdfplumber
    except ImportError:
        raise Exception("PDF support requires: pip install pdfplumber")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                raise ValueError("PDF is empty")

            # Try to extract tables
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    table = max(tables, key=len)
                    if len(table) > 1:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        # Filter empty rows
                        df = df.dropna(how='all').reset_index(drop=True)
                        if len(df) > 0:
                            return df

            raise ValueError("No valid table found in PDF")

    except ValueError as ve:
        raise Exception(f"PDF error: {str(ve)}")
    except Exception as e:
        raise Exception(f"PDF processing error: {str(e)}")
```

---

## Testing Strategy

### Unit Tests
```python
# Test PDF extraction with sample invoices
def test_extract_pdf_simple():
    """Test extraction from simple invoice PDF"""
    # Create test PDF with table
    # Assert extracted columns match expected

def test_extract_pdf_no_table():
    """Test handling of PDF without tables"""
    # Should raise ValueError

def test_extract_pdf_multiple_tables():
    """Test PDF with multiple tables"""
    # Should return the largest table
```

### Integration Tests
```python
def test_load_pdf_mapping():
    """Test full workflow: Load PDF → Drag columns → Save profile"""
    # Load PDF
    # Verify columns appear in left panel
    # Drag to targets
    # Save profile
    # Load profile
    # Assert mapping restored
```

---

## Implementation Checklist

### MVP (Phase 1)
- [ ] Add pdfplumber to requirements.txt
- [ ] Implement `extract_pdf_table()` function
- [ ] Update `load_csv_for_shipment_mapping()` to support PDF
- [ ] Update file dialog filter
- [ ] Update button label "Load Invoice File"
- [ ] Add error handling and messages
- [ ] Test with sample invoice PDFs
- [ ] Update user guide documentation
- [ ] Commit and push to feature branch

### Phase 2 (Future)
- [ ] Multi-table selection UI
- [ ] Column name auto-cleanup
- [ ] Auto-suggest mappings based on content
- [ ] Support for scanned PDF detection
- [ ] OCR option for scanned invoices (pytesseract)

---

## Dependencies

### Install for PDF Support:
```bash
# In your project directory
source venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install pdfplumber
pip freeze > requirements.txt  # Update requirements
```

### Update requirements.txt:
```
pdfplumber>=0.10.0
```

---

## Code Location Reference

| Item | Location |
|------|----------|
| Load function | Line 2681-2702 |
| Setup function | Line 2609-2679 |
| Target mappings | Line 2662-2671 |
| File dialog | Line 2682 |
| Button creation | Line 2632-2634 |
| Drop handler | Line 2704+ |

---

## Recommendations

### Best Practices:
1. **Start with MVP** - Add basic PDF support first
2. **Use pdfplumber** - Most reliable table extraction
3. **Keep UI consistent** - Reuse existing DraggableLabel/DropTarget classes
4. **Error messages** - Be specific about what went wrong
5. **Logging** - Log PDF processing steps for troubleshooting
6. **Testing** - Test with real invoice PDFs from your suppliers

### Why This Approach?
- ✅ Minimal code changes - Reuses existing drag-and-drop UI
- ✅ Consistent with current workflow - No new UI patterns
- ✅ Extensible - Easy to add Phase 2 features later
- ✅ Low risk - PDF support is optional; CSV/Excel still work
- ✅ Professional - Handles errors gracefully

---

## Example Workflow (User Perspective)

1. User clicks "Load Invoice File" button
2. File dialog opens with PDF filter option
3. User selects PDF invoice
4. System extracts table from PDF
5. Columns appear in left panel as DraggableLabel widgets
6. User drags "Part Number" and "Value USD" to required fields
7. User saves profile with supplier name
8. Profile can now be used in Process Shipment tab

**Result:** PDF invoices work seamlessly with existing invoice processing!

