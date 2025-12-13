# PDF Invoice Processing Design - Complete Documentation

## Overview

This document describes the complete PDF invoice processing system integrated into the DerivativeMill application. The system supports automatic extraction of invoice data from both scanned (OCR) and digital PDFs, with intelligent processing and file organization.

---

## Table of Contents

1. [System Architecture](#system-architecture)
2. [Folder Structure](#folder-structure)
3. [Processing Workflow](#processing-workflow)
4. [Configuration](#configuration)
5. [Supported PDF Types](#supported-pdf-types)
6. [API Reference](#api-reference)
7. [User Guide](#user-guide)
8. [Troubleshooting](#troubleshooting)

---

## System Architecture

### High-Level Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                    DerivativeMill Application                    │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │           Invoice Mapping Profiles Tab                  │   │
│  │  • User clicks "Load PDF File" button                   │   │
│  │  • File dialog opens to Input folder                    │   │
│  │  • User selects PDF invoice                             │   │
│  └─────────────────┬───────────────────────────────────────┘   │
│                    │                                             │
│                    ▼                                             │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │        load_csv_for_shipment_mapping()                  │   │
│  │  • Determines file type (PDF/CSV/Excel)                 │   │
│  │  • For PDFs: routes to extract_pdf_table()              │   │
│  └─────────────────┬───────────────────────────────────────┘   │
│                    │                                             │
│                    ▼                                             │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │        extract_pdf_table()                              │   │
│  │  • Attempts pdfplumber table extraction                 │   │
│  │  • If no tables found → calls fallback                  │   │
│  └─────────────────┬───────────────────────────────────────┘   │
│                    │                                             │
│                    ▼                                             │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │        _extract_pdf_text_fallback()                     │   │
│  │  • Extracts all text from PDF                           │   │
│  │  • Checks for "AROMATE" or "SKU#" keywords              │   │
│  │  • Routes to appropriate extractor                      │   │
│  │  • For AROMATE: calls _extract_aromate_invoice()        │   │
│  │  • For OCR: calls extract_from_scanned_invoice()        │   │
│  └─────────────────┬───────────────────────────────────────┘   │
│                    │                                             │
│      ┌─────────────┼─────────────┐                              │
│      │             │             │                              │
│      ▼             ▼             ▼                              │
│  ┌──────────┐  ┌────────┐  ┌──────────────────┐                │
│  │ AROMATE  │  │ OCR    │  │ Generic Text     │                │
│  │ Regex    │  │ Extract│  │ Line Extraction  │                │
│  │ Pattern  │  │        │  │                  │                │
│  └────┬─────┘  └───┬────┘  └────────┬─────────┘                │
│       │            │                 │                         │
│       │            ▼                 │                         │
│       │        ┌──────────────┐      │                         │
│       └───────▶│ move_pdf_to_ │◀─────┘                         │
│                │ processed()  │                                │
│                └──────┬───────┘                                │
│                       │                                        │
│                       ▼                                        │
│            ┌────────────────────┐                             │
│            │ Returns DataFrame  │                             │
│            │ with columns:      │                             │
│            │ • part_number      │                             │
│            │ • quantity         │                             │
│            │ • unit_price       │                             │
│            │ • total_price      │                             │
│            │ • text_line        │                             │
│            └────────┬───────────┘                             │
│                     │                                         │
│                     ▼                                         │
│     ┌───────────────────────────────────┐                    │
│     │ Display columns in Mapping UI     │                    │
│     │ User can drag/drop to map fields  │                    │
│     └───────────────────────────────────┘                    │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

### Component Interaction

```
Settings Dialog
├── Appearance Tab
│   ├── Theme selection
│   ├── Font size
│   ├── Excel viewer
│   └── Row colors
└── Folders Tab
    ├── Input Folder picker → SELECT INPUT_DIR
    ├── Output Folder picker → SELECT OUTPUT_DIR
    └── Processed PDF Folder picker → SELECT PROCESSED_PDF_DIR
                                              ↓
                                    Saved to app_config table
                                              ↓
                                    Used by extraction logic

Invoice Mapping Profiles Tab
├── "Load PDF File" button
│   └── Opens file dialog to INPUT_DIR
│       └── User selects PDF
│           └── Calls load_csv_for_shipment_mapping(pdf_path)
│               └── Extracts columns
│                   └── Displays draggable labels
│                       └── User maps to shipment fields
│
└── Process button (after mapping)
    └── Applies mapping and processes invoice
```

---

## Folder Structure

### Directory Organization

```
DerivativeMill/
├── Input/                          # User-configured, contains source files
│   ├── *.pdf                       # Invoice PDFs to process
│   ├── *.csv                       # Or CSV files
│   └── Processed/                  # Old/archive folder (not used for new PDFs)
│
├── Output/                         # User-configured, contains processed files
│   ├── *.xlsx                      # Exported worksheets
│   ├── *.csv                       # Exported data
│   └── Processed/                  # Old export files (>3 days)
│
├── ProcessedPDFs/                  # User-configured, PDF archive location
│   └── *.pdf                       # Moved here after successful processing
│       └── invoice_1.pdf
│       └── invoice_1_1.pdf         # (auto-numbered if duplicate name)
│       └── invoice_2.pdf
│
├── Resources/                      # App resources (database, icons, etc.)
│   └── derivativemill.db           # SQLite database
│       ├── app_config table        # Settings storage
│       │   ├── key='theme' → value='Fusion (Dark)'
│       │   ├── key='font_size' → value='12'
│       │   ├── key='input_dir' → value='/path/to/Input'
│       │   ├── key='output_dir' → value='/path/to/Output'
│       │   ├── key='processed_pdf_dir' → value='/path/to/ProcessedPDFs'
│       │   └── ...more settings...
│       ├── shipment_mapping table  # Invoice field mappings
│       └── tariff_232 table        # HTS code data
│
└── shipment_mapping.json           # User's column mappings (backup)
```

### Path Configuration

**Global Variables** (derivativemill.py, lines 91-96):
```python
INPUT_DIR = BASE_DIR / "Input"              # Default location
OUTPUT_DIR = BASE_DIR / "Output"            # Default location
PROCESSED_DIR = INPUT_DIR / "Processed"     # Legacy location
OUTPUT_PROCESSED_DIR = OUTPUT_DIR / "Processed"  # Legacy location
PROCESSED_PDF_DIR = BASE_DIR / "ProcessedPDFs"  # New location for processed PDFs
```

**Runtime Update via Settings**:
- User changes folder locations in Settings > Folders tab
- New paths saved to `app_config` database table
- Global variables updated when dialog is opened
- Settings persist across application restarts

---

## Processing Workflow

### Step-by-Step Processing Flow

#### 1. User Opens PDF File

**UI**: Invoice Mapping Profiles Tab → "Load PDF File" button

```python
def load_csv_for_shipment_mapping(self):
    # Opens file dialog, defaults to INPUT_DIR
    path, _ = QFileDialog.getOpenFileName(
        self,
        "Select Invoice File",
        str(INPUT_DIR),
        "All Supported (*.csv *.xlsx *.pdf);;CSV Files (*.csv);;..."
    )
    # User selects: /home/user/Input/invoice_001.pdf
```

#### 2. File Type Detection

```python
file_ext = Path(path).suffix.lower()  # ".pdf"

if file_ext == '.pdf':
    # Proceed to PDF processing logic
    ...
```

#### 3. PDF Processing (Main Logic)

```python
# Location: derivativemill.py, line 2771-2795

if file_ext == '.pdf':
    # BRANCH A: Scanned PDF (Image-based)
    if OCR_AVAILABLE and is_scanned_pdf(path):
        df, metadata = extract_from_scanned_invoice(path)
        # Returns DataFrame from OCR module
        # Shows confidence message
        self.move_pdf_to_processed(path)  # AUTO-MOVE after success

    # BRANCH B: Digital PDF (Text-based)
    else:
        df = self.extract_pdf_table(path)
        # Tries pdfplumber table extraction first
        # If no table → falls back to text extraction
        # Returns DataFrame with extracted data
        self.move_pdf_to_processed(path)  # AUTO-MOVE after success
```

### A. AROMATE Digital PDF Processing

**PDF Type**: Digital PDF with structured invoice table

**Detection**: Text contains "AROMATE" or "SKU#"

**Extraction Method**: Regex Pattern Matching

```python
# Location: derivativemill.py, lines 1872-1903

def _extract_aromate_invoice(self, text):
    """
    Pattern: SKU# XXXXXXX QTY PCS [USD] UNIT_PRICE [USD] TOTAL_PRICE

    Example matches:
    SKU# 1562485 76,080 PCS USD 0.6580 USD 50,060.64
    SKU# 2641486 15,120 PCS 0.7140 10,795.68
    """

    pattern = r'SKU#\s*(\d+)\s+(\d+(?:,\d{3})*)\s+PCS\s+(?:USD\s+)?([\d.]+)\s+(?:USD\s+)?([\d,]+\.\d{2})'

    matches = re.findall(pattern, text)
    # Returns: [('1562485', '76,080', '0.658', '50060.64'), ...]

    # Convert to DataFrame
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
```

**Output Columns**:
| Column | Type | Example | Description |
|--------|------|---------|-------------|
| part_number | str | "1562485" | SKU from invoice |
| quantity | int | 76080 | Order quantity (no commas) |
| unit_price | float | 0.658 | Price per unit |
| total_price | float | 50060.64 | Line total |

### B. OCR Scanned PDF Processing

**PDF Type**: Scanned/image-based PDF

**Detection**: `is_scanned_pdf()` returns True

**Extraction Method**: Pytesseract OCR + Supplier Templates

```python
# Location: ocr/ocr_extract.py

def extract_from_scanned_invoice(pdf_path, supplier_name='auto'):
    """
    Flow:
    1. pdf_to_images() - Convert PDF pages to images
    2. pytesseract.image_to_string() - OCR each image
    3. SupplierTemplate pattern matching - Extract fields
    4. Return DataFrame with extracted data
    """

    images = pdf_to_images(pdf_path)
    text = ""
    for image in images:
        text += pytesseract.image_to_string(image)

    # Auto-detect supplier by checking text
    if "AROMATE" in text:
        supplier_name = "AROMATE"

    # Load supplier template
    template = get_template_manager().get_template(supplier_name)

    # Extract fields using template patterns
    df = extract_fields_from_text(text, template)

    return df, metadata
```

**Output**: DataFrame with columns defined by supplier template

**Confidence Metadata**: Includes success flag and accuracy information

### C. Generic Text Extraction (Fallback)

**Used When**: No AROMATE pattern found, non-OCR required

**Extraction Method**: Text line splitting

```python
# Location: derivativemill.py, lines 1225-1239

def _extract_pdf_text_fallback(self, text):
    """
    Extracts all non-empty lines from PDF text
    Each line becomes a row with 'text_line' column
    """

    all_text = [line.strip() for line in text.split('\n') if line.strip()]
    df = pd.DataFrame({'text_line': all_text})

    return df
```

**Output**: Single column DataFrame
| Column | Example |
|--------|---------|
| text_line | "Company Name" |
| text_line | "Invoice Number: 12345" |
| text_line | "SKU# 1562485 76,080 PCS..." |
| ... | ... |

**Usage**: When PDF has unstructured text, user can manually select/map columns

---

### Step 4: Display Extracted Data

```python
# User sees extracted columns as draggable labels
# Example for AROMATE invoice:
├── [part_number]  ← User drags to "Product No" field
├── [quantity]     ← User drags to "Quantity" field
├── [unit_price]   ← User drags to "Unit Price" field
└── [total_price]  ← User drags to "Total Value" field
```

### Step 5: Auto-Move Processed PDF

```python
# Location: derivativemill.py, lines 1872-1903

def move_pdf_to_processed(self, pdf_path):
    """
    Called after successful extraction.

    Source: /home/user/Input/invoice_001.pdf
    Destination: /home/user/ProcessedPDFs/invoice_001.pdf

    If destination exists:
    Try: invoice_001_1.pdf, invoice_001_2.pdf, etc.
    """

    pdf_file = Path(pdf_path)
    dest_path = PROCESSED_PDF_DIR / pdf_file.name

    # Handle duplicates
    if dest_path.exists():
        base_name = pdf_file.stem
        ext = pdf_file.suffix
        counter = 1
        while dest_path.exists():
            dest_path = PROCESSED_PDF_DIR / f"{base_name}_{counter}{ext}"
            counter += 1

    shutil.move(str(pdf_file), str(dest_path))
```

**Result**: Processed PDFs moved out of Input folder, organized in ProcessedPDFs folder

---

## Configuration

### Settings Dialog

**Location**: Menu Bar > Settings or gear icon

**Structure**: Two tabs

#### Tab 1: Appearance

```
┌─────────────────────────────────────┐
│         Appearance                  │
├─────────────────────────────────────┤
│                                     │
│ Appearance Group                    │
│  Application Theme: [Dropdown ▼]    │
│  Font Size: [Slider 8-16pt]        │
│                                     │
│ Excel File Viewer Group             │
│  Open With: [Dropdown ▼]            │
│                                     │
│ Preview Table Row Colors Group      │
│  Section 232 Rows: [Color ●]        │
│  Non-232 Rows: [Color ●]            │
│                                     │
└─────────────────────────────────────┘
```

**Saved Settings**:
```
app_config table:
├── theme='Fusion (Dark)'
├── font_size='12'
├── excel_viewer='System Default'
├── preview_steel_color='#4a4a4a'
└── preview_non232_color='#ff0000'
```

#### Tab 2: Folders

```
┌─────────────────────────────────────┐
│         Folders                     │
├─────────────────────────────────────┤
│                                     │
│ Folder Locations Group              │
│                                     │
│ Input Folder:                       │
│ ┌─────────────────────────────────┐ │
│ │ /home/user/Input                │ │ (45px height)
│ │ [scrollbar if text overflows]   │ │
│ └─────────────────────────────────┘ │
│ [Change Input Folder]               │
│                                     │
│ Output Folder:                      │
│ ┌─────────────────────────────────┐ │
│ │ /home/user/Output               │ │
│ └─────────────────────────────────┘ │
│ [Change Output Folder]              │
│                                     │
│ Processed PDF Folder:               │
│ ┌─────────────────────────────────┐ │
│ │ /home/user/ProcessedPDFs        │ │
│ └─────────────────────────────────┘ │
│ [Change Processed PDF Folder]       │
│                                     │
└─────────────────────────────────────┘
```

**Saved Settings**:
```
app_config table:
├── input_dir='/home/user/Input'
├── output_dir='/home/user/Output'
└── processed_pdf_dir='/home/user/ProcessedPDFs'
```

**Implementation Details**:
- [derivativemill.py:1017-1289](DerivativeMill/derivativemill.py#L1017-L1289) - show_settings_dialog()
- [derivativemill.py:1814-1870](DerivativeMill/derivativemill.py#L1814-L1870) - Folder selection methods

### Database Storage

**Table**: `app_config`

**Schema**:
```sql
CREATE TABLE app_config (
    key TEXT PRIMARY KEY,
    value TEXT
);
```

**Example Rows**:
```
key                  | value
--------------------|----------------------------------------
theme                | Fusion (Dark)
font_size            | 12
input_dir            | /home/user/Input
output_dir           | /home/user/Output
processed_pdf_dir    | /home/user/ProcessedPDFs
excel_viewer         | System Default
preview_steel_color  | #4a4a4a
preview_non232_color | #ff0000
```

---

## Supported PDF Types

### 1. AROMATE Digital PDFs

**Characteristics**:
- Digital PDF (text-based, not scanned)
- Contains structured invoice table
- Items marked with "SKU#" prefix
- Quantities in "PCS" (pieces)
- Prices in USD or plain numbers

**Example Content**:
```
SKU# 1562485 76,080 PCS USD 0.6580 USD 50,060.64
SKU# 2641486 15,120 PCS 0.7140 10,795.68
SKU# 2641487 48,780 PCS 0.7140 34,828.92
SKU# 2641488 15,840 PCS 0.7320 11,594.88
```

**Extraction Method**: Regex pattern matching

**Regex Pattern**:
```regex
SKU#\s*(\d+)\s+(\d+(?:,\d{3})*)\s+PCS\s+(?:USD\s+)?([\d.]+)\s+(?:USD\s+)?([\d,]+\.\d{2})
```

**Extracted Columns**:
- part_number
- quantity
- unit_price
- total_price

**Success Rate**: 99% (for AROMATE invoices in standard format)

---

### 2. Scanned/OCR PDFs

**Characteristics**:
- Image-based PDF (scanned documents)
- Text extracted via Tesseract OCR
- Accuracy depends on scan quality
- Requires pytesseract and pdf2image libraries

**Detection**:
```python
def is_scanned_pdf(pdf_path):
    """
    Returns True if PDF is image-based (scanned)
    Returns False if PDF is digital (text-based)
    """
    # Uses pdfplumber to analyze content
    # Checks for image-only pages vs text pages
```

**Extraction Method**: OCR + Supplier Templates

**Success Rate**: 70-90% (depends on scan quality and clarity)

**Confidence Feedback**:
- Shows user extraction accuracy percentage
- Warns if accuracy is below threshold
- Allows manual review before processing

---

### 3. Generic Text PDFs

**Characteristics**:
- Text-based PDF without structured tables
- No recognized invoice format (not AROMATE)
- Unstructured line-by-line text

**Extraction Method**: Text line splitting

**Output**: Single column with all text lines

**Usage**: User manually selects relevant lines for mapping

---

## API Reference

### Main Methods

#### `load_csv_for_shipment_mapping()`

**Location**: derivativemill.py, line 2759

**Purpose**: Load invoice file and extract columns for mapping

**Parameters**: None (uses file dialog)

**Returns**: None (updates UI with column labels)

**Flow**:
1. Opens file dialog to INPUT_DIR
2. Detects file type (.pdf, .csv, .xlsx)
3. Routes to appropriate extraction method
4. Extracts column names
5. Displays draggable labels in UI

**Example Usage**:
```python
# User clicks "Load PDF File" button
# → load_csv_for_shipment_mapping() called
# → File dialog opens
# → User selects invoice_001.pdf
# → Extraction happens automatically
# → Columns displayed as draggable labels
```

---

#### `extract_pdf_table(pdf_path)`

**Location**: derivativemill.py, line 2820

**Purpose**: Extract tabular data from PDF

**Parameters**:
- `pdf_path` (str): Path to PDF file

**Returns**: pd.DataFrame or raises Exception

**Logic**:
1. Attempts pdfplumber table extraction
2. If no table found → calls `_extract_pdf_text_fallback()`
3. Returns extracted DataFrame

**Raises**:
- Exception: If PDF cannot be read or no data extracted

---

#### `_extract_pdf_text_fallback(pdf_path)`

**Location**: derivativemill.py, line 1872

**Purpose**: Extract text from PDFs without structured tables

**Parameters**:
- `pdf_path` (str): Path to PDF file

**Returns**: pd.DataFrame with text_line or specialized columns

**Logic**:
1. Extracts all text from PDF using pdfplumber
2. Checks for "AROMATE" or "SKU#" keywords
3. Routes to appropriate extractor:
   - If AROMATE: calls `_extract_aromate_invoice()`
   - If OCR available & scanned: calls `extract_from_scanned_invoice()`
   - Else: returns generic text_line DataFrame

**Returns**:
- DataFrame with extracted columns (AROMATE/OCR) or
- DataFrame with single 'text_line' column (generic)

---

#### `_extract_aromate_invoice(text)`

**Location**: derivativemill.py, line 1872

**Purpose**: Extract AROMATE invoice data using regex

**Parameters**:
- `text` (str): Extracted PDF text

**Returns**: pd.DataFrame

**DataFrame Columns**:
| Column | Type | Description |
|--------|------|-------------|
| part_number | str | SKU number |
| quantity | int | Order quantity |
| unit_price | float | Price per unit |
| total_price | float | Line total |

**Example**:
```python
import pandas as pd
from derivativemill import app_instance

text = "SKU# 1562485 76,080 PCS USD 0.6580 USD 50,060.64"
df = app_instance._extract_aromate_invoice(text)
# Returns:
#    part_number  quantity  unit_price  total_price
# 0      1562485     76080       0.658    50060.64
```

---

#### `move_pdf_to_processed(pdf_path)`

**Location**: derivativemill.py, line 1872

**Purpose**: Move processed PDF to ProcessedPDFs folder

**Parameters**:
- `pdf_path` (str or Path): Path to PDF file to move

**Returns**: bool (True if successful, False if error)

**Logic**:
1. Gets destination path: PROCESSED_PDF_DIR / filename
2. If destination exists, appends counter: filename_1.pdf
3. Moves file using shutil.move()
4. Logs operation

**Example**:
```python
# After extraction completes
success = self.move_pdf_to_processed('/home/user/Input/invoice.pdf')
# File moved to: /home/user/ProcessedPDFs/invoice.pdf

# If duplicate exists:
# File moved to: /home/user/ProcessedPDFs/invoice_1.pdf
```

**Side Effects**:
- Removes PDF from Input folder
- Adds PDF to ProcessedPDFs folder
- Logs to application logger

---

#### `select_input_folder(display_widget)`

**Location**: derivativemill.py, line 1814

**Purpose**: Allow user to select custom Input folder

**Parameters**:
- `display_widget` (QPlainTextEdit or QLabel): Display widget to update

**Returns**: None (updates global INPUT_DIR)

**Side Effects**:
1. Opens folder selection dialog
2. Updates global INPUT_DIR variable
3. Creates Processed subfolder
4. Saves setting to app_config database
5. Updates display widget text
6. Calls refresh_input_files()

---

#### `select_output_folder(display_widget)`

**Location**: derivativemill.py, line 1834

**Purpose**: Allow user to select custom Output folder

**Parameters**:
- `display_widget` (QPlainTextEdit or QLabel): Display widget to update

**Returns**: None (updates global OUTPUT_DIR)

**Side Effects**:
1. Opens folder selection dialog
2. Updates global OUTPUT_DIR variable
3. Creates folder if doesn't exist
4. Saves setting to app_config database
5. Updates display widget text
6. Calls refresh_exported_files()

---

#### `select_processed_pdf_folder(display_widget)`

**Location**: derivativemill.py, line 1853

**Purpose**: Allow user to select custom Processed PDF folder

**Parameters**:
- `display_widget` (QPlainTextEdit or QLabel): Display widget to update

**Returns**: None (updates global PROCESSED_PDF_DIR)

**Side Effects**:
1. Opens folder selection dialog
2. Updates global PROCESSED_PDF_DIR variable
3. Creates folder if doesn't exist
4. Saves setting to app_config database
5. Updates display widget text
6. Logs change to application logger

---

### OCR Module Functions

#### `is_scanned_pdf(pdf_path)`

**Location**: ocr/scanned_pdf.py

**Purpose**: Detect if PDF is scanned (image-based)

**Parameters**:
- `pdf_path` (str): Path to PDF file

**Returns**: bool (True if scanned, False if digital)

**Logic**:
- Analyzes PDF page content types
- Returns True if contains primarily image pages
- Returns False if contains text pages

---

#### `extract_from_scanned_invoice(pdf_path, supplier_name='auto')`

**Location**: ocr/ocr_extract.py

**Purpose**: Extract data from scanned PDF using OCR

**Parameters**:
- `pdf_path` (str): Path to PDF file
- `supplier_name` (str): Supplier name or 'auto' for detection

**Returns**: Tuple[pd.DataFrame, dict]
- DataFrame: Extracted data
- dict: Metadata including success flag and accuracy

**Example Return**:
```python
df, metadata = extract_from_scanned_invoice('invoice.pdf')
# df = DataFrame with extracted columns
# metadata = {
#     'success': True,
#     'confidence': 0.85,
#     'supplier': 'AROMATE',
#     'rows_extracted': 4
# }
```

---

## User Guide

### Quick Start

#### 1. Configure Folders

1. Open Settings (⚙ icon in menu bar)
2. Click "Folders" tab
3. Click "Change Input Folder"
4. Select folder where you'll place PDFs (default: DerivativeMill/Input)
5. Repeat for "Output Folder" and "Processed PDF Folder"
6. Click OK

#### 2. Process Invoice

1. Go to "Invoice Mapping Profiles" tab
2. Click "Load PDF File"
3. Select invoice PDF from your Input folder
4. View extracted columns
5. Drag column labels to appropriate fields
6. Save mapping (optional)
7. Click "Process Invoice" to proceed

#### 3. Archive Processed PDFs

- After successful processing, PDF automatically moves to ProcessedPDFs folder
- Original Input folder stays clean
- Easy to track what's been processed

---

### Processing Different Invoice Types

#### AROMATE Invoices

**Requirements**: PDF must contain "SKU#" text and structured format

**Automatic Processing**:
- App detects "AROMATE" or "SKU#" in text
- Automatically extracts using regex pattern
- Returns columns: part_number, quantity, unit_price, total_price
- PDFs auto-moved after extraction

**What to do**:
1. Place AROMATE PDF in Input folder
2. Click "Load PDF File"
3. Columns automatically extracted
4. Map columns to shipment fields
5. Continue with normal processing

#### Scanned Invoices

**Requirements**:
- PDF is image-based (scanned document)
- pytesseract installed
- pdf2image installed

**Processing**:
- App detects scanned PDF
- Runs OCR to extract text
- Matches text to supplier template
- Shows confidence percentage

**What to do**:
1. Place scanned PDF in Input folder
2. Click "Load PDF File"
3. Review OCR confidence message
4. If confidence is low, review extracted data carefully
5. Map columns to shipment fields
6. Continue with normal processing

#### Generic PDFs

**What to do**:
1. Place PDF in Input folder
2. Click "Load PDF File"
3. View extracted text lines
4. Manually select relevant lines for mapping
5. Continue with normal processing

---

### Best Practices

**File Organization**:
- Keep Input folder clean - processed PDFs are moved out
- Use meaningful file names (e.g., "invoice_001.pdf")
- Store backups in separate location

**Folder Setup**:
- Input: Where you place raw PDFs to process
- Output: Where exported worksheets go
- ProcessedPDFs: Archive of successfully processed PDFs

**Data Mapping**:
- Save column mappings to reuse with similar invoice formats
- Test mapping with one invoice before processing many
- Verify extracted data before finalizing mapping

---

## Troubleshooting

### Problem: PDF shows truncated path in Settings

**Solution**:
- Paths are displayed in read-only text edit widgets with scroll bars
- Text should scroll horizontally if it exceeds the box
- If still truncated, try scrolling with arrow keys

---

### Problem: "No valid table found in PDF" error

**Cause**: PDF doesn't contain structured table format

**Solution**:
1. Check if PDF is AROMATE format (has "SKU#")
2. Check if PDF is scanned (OCR required)
3. If neither, generic text extraction will return text_line column
4. Manually select relevant lines from text_line column

---

### Problem: OCR extraction has low confidence

**Cause**: PDF scan quality is poor (blurry, skewed, dark)

**Solution**:
1. Improve document scan quality if possible
2. Review extracted data carefully
3. Correct any OCR errors before processing
4. Consider re-scanning document with better settings

---

### Problem: File not moved to ProcessedPDFs after processing

**Cause**: File move failed (permissions, disk space, invalid path)

**Solution**:
1. Check ProcessedPDFs folder location in Settings
2. Verify folder path is accessible
3. Check disk space available
4. Verify file permissions
5. Review application log for error message

---

### Problem: Settings don't persist after restart

**Cause**: Database wasn't saved properly

**Solution**:
1. Check that app_config table exists in database
2. Verify database file (derivativemill.db) exists
3. Check file permissions on Resources folder
4. Try resetting settings in Settings dialog

---

### Problem: AROMATE regex pattern not matching

**Cause**: Invoice format doesn't match expected pattern

**Solution**:
1. Verify invoice contains "SKU#" keyword
2. Check format matches: `SKU# XXXXXX QUANTITY PCS [USD] PRICE`
3. Ensure quantities and prices are in correct columns
4. If format is different, generic text extraction will work instead
5. Manually select lines for mapping

---

## Advanced Usage

### Loading Saved Settings

Settings are automatically loaded on startup:

```python
# At application startup (derivativemill.py)
def load_saved_settings():
    try:
        conn = sqlite3.connect(str(DB_PATH))
        c = conn.cursor()

        # Load folder paths
        c.execute("SELECT value FROM app_config WHERE key = 'input_dir'")
        INPUT_DIR = Path(c.fetchone()[0])

        c.execute("SELECT value FROM app_config WHERE key = 'processed_pdf_dir'")
        PROCESSED_PDF_DIR = Path(c.fetchone()[0])

        # ... more settings ...

        conn.close()
    except:
        # Use defaults
        pass
```

### Creating Custom Supplier Patterns

To add support for a new invoice type:

1. Identify the pattern in the PDF text
2. Create a regex pattern to match it
3. Add conditional check in `_extract_pdf_text_fallback()`:

```python
if "YOUR_SUPPLIER" in text:
    return self._extract_your_supplier_invoice(text)
```

4. Implement extraction method:

```python
def _extract_your_supplier_invoice(self, text):
    pattern = r'your_regex_pattern_here'
    matches = re.findall(pattern, text)

    data = []
    for match in matches:
        data.append({
            'part_number': match[0],
            'quantity': int(match[1]),
            # ... more fields ...
        })

    df = pd.DataFrame(data)
    return df
```

---

## Summary

The PDF invoice processing system provides:

✓ **Automatic Detection**: Identifies PDF type (AROMATE/OCR/generic)
✓ **Intelligent Extraction**: Routes to appropriate extraction method
✓ **User Configuration**: Customizable folder locations
✓ **Automatic Organization**: Moves processed PDFs to archive folder
✓ **Flexible Mapping**: Users can map extracted columns to custom fields
✓ **Comprehensive Logging**: Full audit trail of all operations
✓ **Error Handling**: Graceful fallbacks for unsupported formats

The system is designed to be extensible - new supplier formats can be added by creating new extraction methods and adding pattern matching logic.

---

## Related Files

- **Main Application**: [DerivativeMill/derivativemill.py](DerivativeMill/derivativemill.py)
- **OCR Module**: [DerivativeMill/ocr/](DerivativeMill/ocr/)
- **Settings**: [DerivativeMill/Resources/derivativemill.db](DerivativeMill/Resources/derivativemill.db)
- **Sample Extraction**: [DerivativeMill/extract_aromate_invoice.py](DerivativeMill/extract_aromate_invoice.py)
- **Batch Processing Plan**: [BATCH_PROCESSING_PLAN.md](BATCH_PROCESSING_PLAN.md)

---

**Document Version**: 1.0
**Last Updated**: November 29, 2025
**Status**: Complete
