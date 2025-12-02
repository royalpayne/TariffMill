# Batch PDF Invoice Processing Workflow - Implementation Plan

## Overview
Redesign the invoice extraction workflow to support batch processing of PDF documents from supplier-specific folders, with output delivered to a user-configured folder.

## Current Architecture
- **Input Folder**: `DerivativeMill/Input` (configurable in Settings)
- **Output Folder**: `DerivativeMill/Output` (configurable in Settings)
- **Tabs**:
  - Process Shipment (single invoice processing)
  - Invoice Mapping Profiles (supplier mapping configurations)
  - Parts Import
  - Parts View
  - Log View
  - Customs Config
  - User Guide

## Proposed Workflow

### 1. Folder Structure
```
Input/
├── AROMATE/
│   ├── invoice_1.pdf
│   ├── invoice_2.pdf
│   └── ...
├── [Other Supplier Name]/
│   ├── invoice_1.pdf
│   └── ...
└── ...

Output/
├── invoice_1_extracted_2025-01-15_14-30-22.csv
├── invoice_2_extracted_2025-01-15_14-30-23.csv
├── other_invoice_extracted_2025-01-15_14-30-24.csv
└── ...
```

**Note**: Each PDF is extracted individually and saved as a separate CSV file with its own timestamp.

### 2. User Interface Changes

#### Settings Dialog - Add Supplier Folder Management
- **Current**: Shows Input/Output folder paths
- **New**: Add "Supplier PDF Folders" section
  - Shows list of supplier folders in Input directory
  - Button to "Manage Supplier Folders"
  - Can add/remove supplier folders
  - Automatically detects supplier-named subfolders in Input directory

#### New Tab: "Batch Invoice Processing" (optional) OR add to Process Shipment tab
**Option A: New Tab** - Creates dedicated batch processing interface
**Option B: Add to Process Shipment** - Add batch processing section to existing tab (RECOMMENDED for simplicity)

**Recommended: Add to existing "Process Shipment" tab**
- Add "Batch Processing" group below "Actions" group
- "Process All PDFs from Folder" button
- Dropdown to select which supplier folder to process
- Progress indicator
- Results log showing:
  - Files processed
  - Files skipped (if extraction failed)
  - Output file created

### 3. Implementation Components

#### A. Supplier Folder Detection
**File**: `derivativemill.py` (new method)
```python
def get_supplier_folders():
    """Scan INPUT_DIR for supplier-named subfolders"""
    suppliers = []
    for folder in INPUT_DIR.iterdir():
        if folder.is_dir() and folder.name != "Processed":
            suppliers.append(folder.name)
    return sorted(suppliers)
```

#### B. Batch PDF Processing
**File**: `derivativemill.py` (new method)
```python
def process_supplier_pdfs(supplier_name):
    """
    Process all PDFs in supplier folder - extract each PDF individually.

    Flow:
    1. Get folder: INPUT_DIR / supplier_name
    2. Find all .pdf files
    3. For each PDF:
       - Extract using _extract_aromate_invoice() or appropriate extractor
       - Save extracted data to individual CSV file with timestamp
       - Track success/error for this PDF
    4. Move processed PDF to Processed folder
    5. Return results and status with list of all output files
    """
    supplier_folder = INPUT_DIR / supplier_name
    output_files = []
    errors = []

    pdf_files = list(supplier_folder.glob("*.pdf"))

    for pdf_file in pdf_files:
        try:
            df = extract_pdf_data(pdf_file)  # Use existing extraction

            # Save extracted data to individual CSV file
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            pdf_name = pdf_file.stem  # filename without .pdf extension
            output_file = OUTPUT_DIR / f"{pdf_name}_extracted_{timestamp}.csv"
            df.to_csv(output_file, index=False)
            output_files.append(output_file)

            # Move processed PDF to Processed folder
            self.move_pdf_to_processed(pdf_file)

        except Exception as e:
            errors.append((pdf_file.name, str(e)))

    return {
        'success': len(output_files) > 0,
        'output_files': output_files,
        'files_processed': len(output_files),
        'files_skipped': len(errors),
        'errors': errors
    }
```

#### C. UI Components

**1. Add to Settings Dialog** (in show_settings_dialog method)
```
New "Supplier Folders" group:
- List widget showing supplier folders in Input directory
- "Auto-Detect Suppliers" button (scans Input folder)
- "Add Supplier Folder" button (creates new subfolder)
- "Remove Selected" button
- "Open Folder" button (opens in file explorer)
```

**2. Add to Process Shipment Tab** (in setup_process_tab method)
```
New "Batch Processing" group (after Actions group):
- Dropdown: "Select Supplier to Process"
- "Process All PDFs" button (green/primary style)
- Progress bar (shows file count during processing)
- Status text:
  - "Processing: file 3 of 5..."
  - "Complete: 5 files processed, output saved to [filename]"
  - "Error: Skipped 2 files due to extraction errors"
```

### 4. Integration Points

#### Extract PDF Data Function (unified)
Create a single extraction dispatcher that:
1. Detects PDF type (AROMATE, other suppliers, generic)
2. Routes to appropriate extraction method
3. Returns standardized DataFrame

**File**: `derivativemill.py` (modify extract_pdf_table method)
```python
def extract_pdf_data(pdf_path):
    """Unified extraction for single PDF, used by both single and batch processing"""
    # Try table extraction first
    df = self.extract_pdf_table(pdf_path)
    if df is not None:
        return df

    # Fallback to text extraction (AROMATE detection already in place)
    return self._extract_pdf_text_fallback(pdf_path)
```

### 5. Data Flow

```
User Actions:
1. Settings: Creates/manages supplier folders in Input/
2. Puts PDF files in Input/[Supplier Name]/ folders
3. Process Shipment tab: Selects supplier from dropdown
4. Clicks "Process All PDFs"
        ↓
Batch Processing Flow:
5. Scans Input/[Supplier Name]/ for all .pdf files
6. For each PDF:
   - Extract data using appropriate method
   - Save extracted data to individual CSV file
   - Move PDF to Processed folder
   - Track success/error for this PDF
7. Update UI with results:
   - Show files processed count
   - Show any errors/skipped files
   - List all output files created
```

### 6. Settings Dialog Changes

**New "Supplier Folder Management" Group**:
```
┌─────────────────────────────────────┐
│ Supplier Folder Management          │
├─────────────────────────────────────┤
│ Suppliers in Input Folder:          │
│ [AROMATE           ] [Open]         │
│ [Supplier B        ] [Open]         │
│ [Supplier C        ] [Open]         │
│                                     │
│ [+ Add Supplier]  [- Remove]        │
│ [Auto-Detect]                       │
├─────────────────────────────────────┤
│ info: Organize PDFs in folders by   │
│       supplier name. Then batch     │
│       process them all at once.     │
└─────────────────────────────────────┘
```

### 7. Process Shipment Tab Changes

**Add After "Actions" Group**:
```
┌──────────────────────────────────────┐
│ Batch Invoice Processing             │
├──────────────────────────────────────┤
│ Select Supplier: [Dropdown ▼]       │
│                                      │
│ [Process All PDFs]                  │
│                                      │
│ Status: [Progress bar 0%  ]          │
│ Processing: 3 of 5 files...          │
│                                      │
│ Last Result:                         │
│ ✓ 5 files processed                  │
│ ✓ Output files created:              │
│   - invoice_1_extracted_...csv       │
│   - invoice_2_extracted_...csv       │
│   - invoice_3_extracted_...csv       │
│   - invoice_4_extracted_...csv       │
│   - invoice_5_extracted_...csv       │
│ ⚠ 0 files skipped                    │
└──────────────────────────────────────┘
```

### 8. File Changes Required

1. **derivativemill.py**:
   - `get_supplier_folders()` - detect supplier folders
   - `process_supplier_pdfs(supplier_name)` - batch processing logic
   - `extract_pdf_data(pdf_path)` - unified extraction
   - Add "Batch Processing" group to `setup_process_tab()`
   - Add "Supplier Folder Management" group to `show_settings_dialog()`

2. **extract_aromate_invoice.py** (existing standalone tool):
   - Can remain as-is for manual testing
   - Not required for app integration

3. **ocr module**:
   - Already has supplier template system
   - No changes needed for batch processing

### 9. Execution Flow

**Batch Processing Button Click**:
```
1. Get selected supplier name from dropdown
2. Get supplier folder path: INPUT_DIR / supplier_name
3. Scan for *.pdf files
4. Show progress dialog with file count
5. For each PDF:
   a. Extract data (with error handling)
   b. Save extracted data to individual CSV file with timestamp
   c. Move PDF to Processed folder
   d. Update progress
   e. Log result
6. Show completion dialog:
   - Total files processed count
   - List of all output CSV files created
   - Any errors/skipped files with reasons
7. Update Exported Files list with all new output files
8. Update batch processing status display with results
```

### 10. Error Handling

- **No PDFs found**: Show message "No PDF files found in [Supplier] folder"
- **Extraction errors**: Skip file, log error, continue processing
- **Empty results**: Show message "No data could be extracted from any PDF"
- **File write errors**: Show error dialog with file path
- **Invalid supplier folder**: Show "Supplier folder not found" error

### 11. Benefits of This Approach

✓ **User-Friendly**: Intuitive folder-based organization by supplier
✓ **Scalable**: Process any number of files at once
✓ **Consistent**: Uses existing extraction methods (already tested)
✓ **Non-Breaking**: Doesn't change existing single-file workflow
✓ **Configurable**: Settings panel for managing supplier folders
✓ **Integrated**: Uses existing Input/Output paths from settings

### 12. Future Enhancements

- Support for multiple extraction formats per supplier (regex patterns)
- Scheduled/automated batch processing
- Email notifications on completion
- Archive processed PDFs to Input/Processed folder
- Excel export option for batch results
- Merge multiple batches into single report

---

## Recommendation

Implement Steps 1-8 first (core batch processing functionality), then iterate on UI polish and error handling based on user feedback.

**Priority Order**:
1. Add batch processing methods to derivativemill.py
2. Add "Batch Processing" group to Process Shipment tab UI
3. Add "Supplier Folder Management" to Settings dialog
4. Test with sample PDFs
5. Add error handling and logging
6. Document workflow for users
