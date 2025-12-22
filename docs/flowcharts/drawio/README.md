# TariffMill Draw.io Flowcharts

This folder contains professional Draw.io (diagrams.net) flowchart files for TariffMill documentation.

## Files

| File | Description |
|------|-------------|
| `01_invoice_processing_workflow.drawio` | Complete invoice processing flow from file upload to export |
| `02_parts_master_data_flow.drawio` | Parts Master database data flow and management |
| `03_ocrmill_template_processing.drawio` | OCRMill template matching and data extraction |
| `04_section_232_301_tariff_detection.drawio` | Material classification and tariff detection logic |
| `05_application_architecture.drawio` | System architecture, components, and database schema |
| `06_user_workflow.drawio` | Complete end-to-end user journey with all phases |

## How to View/Edit

### Option 1: diagrams.net (Recommended)
1. Go to [diagrams.net](https://app.diagrams.net/)
2. Click "Open Existing Diagram"
3. Select the `.drawio` file from this folder
4. Edit and export as needed

### Option 2: VS Code Extension
1. Install the "Draw.io Integration" extension in VS Code
2. Open any `.drawio` file directly in VS Code
3. Edit with the built-in editor

### Option 3: Desktop Application
1. Download Draw.io Desktop from [GitHub Releases](https://github.com/jgraph/drawio-desktop/releases)
2. Open `.drawio` files directly

## Export Options

From diagrams.net, you can export to:
- **PNG/JPEG** - For documentation and presentations
- **SVG** - Scalable vector graphics for web
- **PDF** - For printing and sharing
- **HTML** - Interactive web pages
- **Visio (VSDX)** - For Microsoft Visio users

## Color Legend

All flowcharts use consistent color coding:

| Color | Meaning |
|-------|---------|
| Green (#d5e8d4) | Success, completion, output |
| Blue (#dae8fc) | Process steps, actions |
| Yellow (#fff2cc) | Decisions, warnings |
| Purple (#e1d5e7) | User input, templates |
| Red (#f8cecc) | Errors, fixes, edits |
| Gray (#f5f5f5) | Information, reference |

## Material Type Colors (Tariff Detection)

| Color | Material |
|-------|----------|
| Blue (#0000FF) | Steel (25% tariff) |
| Green (#00AA00) | Aluminum (10% tariff) |
| Orange (#FF8C00) | Copper |
| Brown (#8B4513) | Wood |
| Purple (#800080) | Automotive |

## Version

Documentation updated for TariffMill v0.94.0
