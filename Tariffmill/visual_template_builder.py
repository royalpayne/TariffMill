"""
Visual Template Builder for OCRMill
Text-based field definition workflow for template creation.
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QPlainTextEdit, QLineEdit, QGroupBox, QFormLayout,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QSplitter, QWidget, QFileDialog, QMessageBox,
    QScrollArea, QFrame, QSpinBox, QListWidget, QListWidgetItem,
    QMenu, QAction, QToolBar, QStatusBar, QToolButton, QButtonGroup,
    QRadioButton, QCheckBox, QSlider, QSizePolicy, QTextEdit,
    QInputDialog, QTabWidget
)
from PyQt5.QtCore import Qt, pyqtSignal, QRect, QPoint, QSize, QRectF
from PyQt5.QtGui import (
    QFont, QPixmap, QPainter, QPen, QColor, QBrush, QImage,
    QCursor, QKeySequence, QTextCharFormat, QTextCursor, QSyntaxHighlighter
)

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


class DetectedPattern:
    """Represents a detected pattern in the text."""

    def __init__(self, pattern_type: str, pattern: str, sample_matches: List[str],
                 description: str, confidence: float = 0.5):
        self.pattern_type = pattern_type  # 'line_item', 'invoice_number', 'po_number', etc.
        self.pattern = pattern
        self.sample_matches = sample_matches
        self.description = description
        self.confidence = confidence


class PatternDetector:
    """Detects common patterns in invoice text with table structure recognition."""

    # Common line item patterns - ordered by specificity (most specific first)
    LINE_ITEM_PATTERNS = [
        # HTS code items: "HTS#1234567890-DESCRIPTION 10 PCS $100.00"
        (r'HTS#(\d{10})-([A-Z\s]+?)\s+(\d+)\s*(?:PCS?|UNITS?)?\s+\$?([\d,]+\.?\d*)',
         'hts_desc_qty_price', 'HTS# Description Qty Price', ['hs_code', 'description', 'quantity', 'total_price']),

        # PO + Part + Qty + Price: "40049557 21-250464 315 $686.70"
        (r'^(\d{8})\s+(\d{2}-\d{6})\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)',
         'po_part_qty_price', 'PO# Part# Qty Price', ['po_number', 'part_number', 'quantity', 'total_price']),

        # Item Code + Qty + Price (simple): "18-123456 100 25.50"
        (r'^(\d{2}-\d{6})\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)',
         'code_qty_price', 'ItemCode Qty Price', ['part_number', 'quantity', 'total_price']),

        # Part + Qty + Unit + Total (4 columns): "ABC-123 10 $5.00 $50.00"
        (r'^([A-Z0-9][\w\-\.]+)\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)\s+\$?([\d,]+\.?\d*)',
         'part_qty_unit_total', 'Part# Qty UnitPrice Total', ['part_number', 'quantity', 'unit_price', 'total_price']),

        # Part + Description + Qty + Price: "ABC-123 Widget Assembly 10 $50.00"
        (r'^([A-Z0-9][\w\-\.]+)\s+([A-Za-z][\w\s]{3,30}?)\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)',
         'part_desc_qty_price', 'Part# Description Qty Price', ['part_number', 'description', 'quantity', 'total_price']),

        # Generic: Number + Text + Number + Currency
        (r'^(\d+)\s+([A-Za-z][\w\s\-]{3,40}?)\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)',
         'num_text_qty_price', 'LineNo Description Qty Price', ['line_number', 'description', 'quantity', 'total_price']),

        # Chinese/international format with currency symbols: "ABC123 100 PCS USD 50.00"
        (r'^([A-Z0-9][\w\-\.]+)\s+(\d+(?:[.,]\d+)?)\s*(?:PCS?|UNITS?|KG|M|EA)?\s+(?:USD|EUR|CNY|RMB)?\s*\$?([\d,]+\.?\d*)',
         'part_qty_currency_price', 'Part# Qty Price', ['part_number', 'quantity', 'total_price']),

        # Description-first format: "Widget Assembly ABC-123 10 $50.00"
        (r'^([A-Za-z][\w\s]{5,40}?)\s+([A-Z0-9][\w\-\.]+)\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)',
         'desc_part_qty_price', 'Description Part# Qty Price', ['description', 'part_number', 'quantity', 'total_price']),
    ]

    # Invoice number patterns
    INVOICE_PATTERNS = [
        (r'[Ii]nvoice\s*(?:#|[Nn]o\.?|[Nn]umber)?\s*:?\s*([A-Z0-9][\w\-/]+)', 'Invoice #'),
        (r'[Ii]nv\.?\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-/]+)', 'Inv #'),
        (r'[Bb]ill\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-/]+)', 'Bill #'),
        (r'[Dd]ocument\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-/]+)', 'Document #'),
        (r'[Rr]eference\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-/]+)', 'Reference #'),
    ]

    # PO/Project number patterns
    PO_PATTERNS = [
        (r'[Pp]\.?[Oo]\.?\s*(?:#|[Nn]o\.?)?\s*:?\s*(\d{6,})', 'PO #'),
        (r'[Pp]urchase\s+[Oo]rder\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-]+)', 'Purchase Order'),
        (r'[Pp]roject\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-]+)', 'Project #'),
        (r'[Oo]rder\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-]+)', 'Order #'),
        (r'[Jj]ob\s*(?:#|[Nn]o\.?)?\s*:?\s*([A-Z0-9][\w\-]+)', 'Job #'),
        (r'\b(400\d{5})\b', 'Sigma PO'),  # Sigma-style PO numbers
        (r'\b(US\d{2}[A-Z]\d{4}[a-z]?)\b', 'US Project Code'),  # mmcite-style project codes
    ]

    # Table header keywords to detect table structure
    TABLE_HEADERS = {
        'part': ['part', 'item', 'sku', 'code', 'article', 'art', 'product'],
        'description': ['description', 'desc', 'name', 'details'],
        'quantity': ['qty', 'quantity', 'units', 'pcs', 'pieces', 'amount'],
        'price': ['price', 'amount', 'total', 'value', 'cost', 'rate'],
        'unit_price': ['unit price', 'unit cost', 'rate', 'each', 'per unit'],
    }

    @classmethod
    def detect_patterns(cls, text: str) -> List[DetectedPattern]:
        """Detect all patterns in the text."""
        detected = []

        # First try table structure detection
        table_patterns = cls._detect_table_structure(text)
        detected.extend(table_patterns)

        # Then try predefined line item patterns
        line_item_results = cls._detect_line_items(text)
        detected.extend(line_item_results)

        # Detect invoice numbers
        invoice_results = cls._detect_field_pattern(text, cls.INVOICE_PATTERNS, 'invoice_number')
        detected.extend(invoice_results)

        # Detect PO/Project numbers
        po_results = cls._detect_field_pattern(text, cls.PO_PATTERNS, 'project_number')
        detected.extend(po_results)

        # Detect company/supplier indicators
        indicator_results = cls._detect_company_indicators(text)
        detected.extend(indicator_results)

        return detected

    @classmethod
    def _detect_table_structure(cls, text: str) -> List[DetectedPattern]:
        """Detect table structure by finding header rows and matching data rows."""
        results = []
        lines = text.split('\n')

        # Look for header rows with multiple column keywords
        for i, line in enumerate(lines):
            line_lower = line.lower().strip()
            if not line_lower:
                continue

            # Count how many header keywords appear in this line
            header_cols = {}
            for col_type, keywords in cls.TABLE_HEADERS.items():
                for keyword in keywords:
                    if keyword in line_lower:
                        header_cols[col_type] = keyword
                        break

            # If we found at least 2 column headers, this might be a table header
            if len(header_cols) >= 2:
                # Try to build a pattern from the following lines
                data_lines = []
                for j in range(i + 1, min(i + 20, len(lines))):
                    data_line = lines[j].strip()
                    if not data_line:
                        continue
                    # Skip if it looks like another header or footer
                    if any(kw in data_line.lower() for kw in ['total', 'subtotal', 'tax', 'page']):
                        break
                    # Check if line has numeric content (likely a data row)
                    if re.search(r'\d+[.,]?\d*', data_line):
                        data_lines.append(data_line[:100])

                if len(data_lines) >= 2:
                    # Generate a pattern based on the data structure
                    pattern = cls._generate_pattern_from_samples(data_lines)
                    if pattern:
                        results.append(DetectedPattern(
                            pattern_type='line_item',
                            pattern=pattern,
                            sample_matches=data_lines[:5],
                            description=f"Table: {', '.join(header_cols.keys())} ({len(data_lines)} rows)",
                            confidence=min(1.0, 0.5 + len(data_lines) / 15)
                        ))

        return results[:2]  # Return top 2 table patterns

    @classmethod
    def _generate_pattern_from_samples(cls, samples: List[str]) -> Optional[str]:
        """Generate a regex pattern by analyzing sample data rows."""
        if len(samples) < 2:
            return None

        # Analyze the structure of sample lines
        # Look for common patterns: alphanumeric codes, numbers, currency amounts

        # Try to identify column structure
        patterns_tried = []

        # Pattern: Starts with code (alphanumeric with possible dashes)
        code_start = r'^([A-Z0-9][\w\-\.]{2,20})'

        # Common column patterns
        number_col = r'(\d+(?:[.,]\d+)?)'
        currency_col = r'\$?([\d,]+\.?\d*)'
        text_col = r'([A-Za-z][\w\s]{3,40}?)'

        # Try various combinations
        pattern_options = [
            code_start + r'\s+' + number_col + r'\s+' + currency_col + r'\s+' + currency_col,  # Code Qty Unit Total
            code_start + r'\s+' + text_col + r'\s+' + number_col + r'\s+' + currency_col,  # Code Desc Qty Price
            code_start + r'\s+' + number_col + r'\s+' + currency_col,  # Code Qty Price
            r'^' + number_col + r'\s+' + text_col + r'\s+' + number_col + r'\s+' + currency_col,  # No Desc Qty Price
        ]

        best_pattern = None
        best_matches = 0

        for pattern in pattern_options:
            try:
                compiled = re.compile(pattern, re.IGNORECASE)
                matches = sum(1 for s in samples if compiled.match(s))
                if matches > best_matches:
                    best_matches = matches
                    best_pattern = pattern
            except re.error:
                continue

        # Only return pattern if it matches at least half the samples
        if best_pattern and best_matches >= len(samples) / 2:
            return best_pattern

        return None

    @classmethod
    def _detect_line_items(cls, text: str) -> List[DetectedPattern]:
        """Detect line item patterns using predefined patterns."""
        results = []
        lines = text.split('\n')

        for pattern, name, description, field_names in cls.LINE_ITEM_PATTERNS:
            matches = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # Skip header-like lines
                if any(h in line.lower() for h in ['description', 'quantity', 'price', 'total', 'item no']):
                    continue
                match = re.match(pattern, line, re.IGNORECASE)
                if match:
                    matches.append(line[:80])  # First 80 chars

            if len(matches) >= 2:  # Need at least 2 matches to be confident
                confidence = min(1.0, len(matches) / 10 + 0.3)
                results.append(DetectedPattern(
                    pattern_type='line_item',
                    pattern=pattern,
                    sample_matches=matches[:5],  # Show up to 5 samples
                    description=f"{description} ({len(matches)} matches)",
                    confidence=confidence
                ))

        # Sort by number of matches (confidence)
        results.sort(key=lambda x: x.confidence, reverse=True)
        return results[:3]  # Return top 3 patterns

    @classmethod
    def _detect_field_pattern(cls, text: str, patterns: List[Tuple], field_type: str) -> List[DetectedPattern]:
        """Detect field patterns like invoice numbers."""
        results = []

        for pattern, description in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
            if matches:
                # Get unique matches
                unique_matches = list(set(matches[:10]))
                confidence = min(1.0, 0.5 + len(unique_matches) * 0.1)
                results.append(DetectedPattern(
                    pattern_type=field_type,
                    pattern=pattern,
                    sample_matches=unique_matches[:5],
                    description=f"{description}: {unique_matches[0]}",
                    confidence=confidence
                ))

        # Sort by confidence
        results.sort(key=lambda x: x.confidence, reverse=True)
        return results[:2]  # Return top 2 patterns

    @classmethod
    def _detect_company_indicators(cls, text: str) -> List[DetectedPattern]:
        """Detect company/supplier name indicators for can_process()."""
        results = []

        # Look for company name patterns
        company_patterns = [
            (r'(?:From|Supplier|Seller|Vendor|Shipper|Exporter)[\s:]+([A-Z][A-Za-z\s&\.]+(?:Ltd|LLC|Inc|Corp|Co\.?|GmbH|Limited|CORPORATION|s\.r\.o\.))', 'Supplier'),
            (r'^([A-Z][A-Z\s&\.]{5,}(?:LTD|LLC|INC|CORP|CO\.?|GMBH|LIMITED|CORPORATION|S\.R\.O\.))\s*$', 'Company Header'),
            (r'(?:Company|Business)[\s:]+([A-Z][A-Za-z\s&\.]+)', 'Company Name'),
        ]

        for pattern, description in company_patterns:
            matches = re.findall(pattern, text, re.MULTILINE)
            if matches:
                # Clean up matches
                unique = list(set(m.strip() for m in matches if len(m.strip()) > 5))
                if unique:
                    results.append(DetectedPattern(
                        pattern_type='indicator',
                        pattern=pattern,
                        sample_matches=unique[:3],
                        description=f"{description}: {unique[0][:30]}",
                        confidence=0.7
                    ))

        return results[:2]


class DefinedField:
    """Represents a field defined from text selection."""

    def __init__(self, name: str, field_type: str, sample_value: str,
                 context_before: str = "", context_after: str = ""):
        self.name = name
        self.field_type = field_type  # invoice_number, project_number, manufacturer, line_item, indicator
        self.sample_value = sample_value
        self.context_before = context_before
        self.context_after = context_after
        self.pattern = ""  # Auto-generated or user-modified regex
        self.color = self._get_color_for_type(field_type)

    def _get_color_for_type(self, field_type: str) -> QColor:
        """Get color for field type."""
        colors = {
            'invoice_number': QColor(46, 204, 113, 150),   # Green
            'project_number': QColor(155, 89, 182, 150),   # Purple
            'manufacturer': QColor(52, 152, 219, 150),     # Blue
            'line_item': QColor(230, 126, 34, 150),        # Orange
            'indicator': QColor(241, 196, 15, 150),        # Yellow
            'text': QColor(149, 165, 166, 150),            # Gray
        }
        return colors.get(field_type, colors['text'])


class TextTableParser:
    """Parse PDF text into a table structure based on spacing/alignment."""

    @staticmethod
    def parse_by_whitespace(text: str, min_spaces: int = 2) -> List[List[str]]:
        """
        Parse text by splitting on multiple whitespace characters.
        This is the most reliable method for invoice data.
        """
        lines = text.split('\n')
        rows = []

        for line in lines:
            if not line.strip():
                rows.append([''])
                continue

            # Split by 2+ spaces (configurable)
            pattern = r'\s{' + str(min_spaces) + r',}'
            parts = re.split(pattern, line.strip())
            # Filter out empty parts
            parts = [p.strip() for p in parts if p.strip()]
            rows.append(parts if parts else [''])

        return rows

    @staticmethod
    def parse_fixed_width(text: str, num_columns: int = 5) -> List[List[str]]:
        """
        Parse text into fixed number of columns based on spacing.
        """
        lines = text.split('\n')
        rows = []

        for line in lines:
            if not line.strip():
                rows.append([''] * num_columns)
                continue

            # Split by 2+ spaces
            parts = re.split(r'\s{2,}', line.strip())
            parts = [p.strip() for p in parts if p.strip()]

            # Pad or trim to exact column count
            if len(parts) < num_columns:
                parts.extend([''] * (num_columns - len(parts)))
            elif len(parts) > num_columns:
                # Merge extra parts into last column
                parts = parts[:num_columns-1] + [' '.join(parts[num_columns-1:])]

            rows.append(parts)

        return rows

    @staticmethod
    def detect_optimal_columns(text: str) -> int:
        """
        Analyze text to find the optimal number of columns.
        """
        rows = TextTableParser.parse_by_whitespace(text)
        if not rows:
            return 1

        # Count column occurrences
        col_counts = {}
        for row in rows:
            n = len(row)
            if n > 0:
                col_counts[n] = col_counts.get(n, 0) + 1

        if not col_counts:
            return 1

        # Find the most common column count (excluding 1-column rows which are often headers/text)
        filtered = {k: v for k, v in col_counts.items() if k > 1}
        if filtered:
            return max(filtered, key=filtered.get)

        return max(col_counts, key=col_counts.get)


class TableTextViewer(QWidget):
    """Widget for displaying text in table format with selection support."""

    field_add_requested = pyqtSignal(str, int, int)  # value, row, col

    def __init__(self, parent=None):
        super().__init__(parent)
        self.full_text = ""
        self.rows_data = []
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        # Controls row 1
        controls1 = QHBoxLayout()

        controls1.addWidget(QLabel("Columns:"))
        self.col_spin = QSpinBox()
        self.col_spin.setRange(1, 15)
        self.col_spin.setValue(6)
        self.col_spin.setToolTip("Number of columns to display")
        self.col_spin.valueChanged.connect(self.refresh_table)
        controls1.addWidget(self.col_spin)

        self.auto_detect_btn = QPushButton("Auto-Detect")
        self.auto_detect_btn.setToolTip("Automatically detect optimal column count")
        self.auto_detect_btn.clicked.connect(self.auto_detect_columns)
        controls1.addWidget(self.auto_detect_btn)

        controls1.addWidget(QLabel("  Min Spaces:"))
        self.space_spin = QSpinBox()
        self.space_spin.setRange(1, 5)
        self.space_spin.setValue(2)
        self.space_spin.setToolTip("Minimum spaces to treat as column separator")
        self.space_spin.valueChanged.connect(self.refresh_table)
        controls1.addWidget(self.space_spin)

        controls1.addStretch()

        self.info_label = QLabel("")
        self.info_label.setStyleSheet("color: #888;")
        controls1.addWidget(self.info_label)

        layout.addLayout(controls1)

        # Selection info and action row
        selection_row = QHBoxLayout()

        self.selection_label = QLabel("Click a cell to select, double-click to add as field")
        self.selection_label.setStyleSheet("color: #666; font-style: italic;")
        selection_row.addWidget(self.selection_label)

        selection_row.addStretch()

        self.btn_add_field = QPushButton("+ Add Selected as Field")
        self.btn_add_field.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 5px 10px;")
        self.btn_add_field.setEnabled(False)
        self.btn_add_field.clicked.connect(self.add_selected_as_field)
        selection_row.addWidget(self.btn_add_field)

        layout.addLayout(selection_row)

        # Table widget
        self.table = QTableWidget()
        self.table.setFont(QFont("Consolas", 9))
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.cellDoubleClicked.connect(self.on_cell_double_clicked)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(22)  # Compact rows
        layout.addWidget(self.table)

        # Store selected cell info
        self.selected_value = ""
        self.selected_row = -1
        self.selected_col = -1

    def set_text(self, text: str):
        """Set text and parse into table."""
        self.full_text = text
        self.auto_detect_columns()

    def auto_detect_columns(self):
        """Auto-detect column structure."""
        if not self.full_text:
            return

        # Detect optimal column count
        optimal = TextTableParser.detect_optimal_columns(self.full_text)
        self.col_spin.blockSignals(True)
        self.col_spin.setValue(max(optimal, 1))
        self.col_spin.blockSignals(False)
        self.refresh_table()

    def refresh_table(self):
        """Refresh table with current settings."""
        if not self.full_text:
            return

        num_cols = self.col_spin.value()
        min_spaces = self.space_spin.value()

        # Parse with current settings
        raw_rows = TextTableParser.parse_by_whitespace(self.full_text, min_spaces)

        # Normalize to fixed columns
        self.rows_data = []
        for row in raw_rows:
            if len(row) < num_cols:
                row = row + [''] * (num_cols - len(row))
            elif len(row) > num_cols:
                row = row[:num_cols-1] + [' '.join(row[num_cols-1:])]
            self.rows_data.append(row)

        self._populate_table(self.rows_data, num_cols)

    def _populate_table(self, rows: List[List[str]], num_cols: int):
        """Populate the table widget with parsed rows."""
        self.table.clear()
        self.table.setRowCount(len(rows))
        self.table.setColumnCount(num_cols)

        # Set headers with guessed names
        headers = self._guess_headers(rows, num_cols)
        self.table.setHorizontalHeaderLabels(headers)

        # Populate cells
        for row_idx, row in enumerate(rows):
            for col_idx, cell in enumerate(row):
                if col_idx < num_cols:
                    item = QTableWidgetItem(cell)
                    # Highlight cells that look like data values
                    if cell and self._is_data_value(cell):
                        item.setBackground(QColor(240, 248, 255))  # Light blue
                    self.table.setItem(row_idx, col_idx, item)

        # Resize columns to content with max width
        self.table.resizeColumnsToContents()
        for i in range(num_cols):
            if self.table.columnWidth(i) > 300:
                self.table.setColumnWidth(i, 300)

        # Update info
        non_empty_rows = sum(1 for row in rows if any(cell for cell in row))
        self.info_label.setText(f"{non_empty_rows} rows × {num_cols} cols")

    def _guess_headers(self, rows: List[List[str]], num_cols: int) -> List[str]:
        """Try to guess column headers based on content."""
        headers = []
        for i in range(num_cols):
            # Check first few rows for header-like content
            col_values = [rows[r][i] for r in range(min(10, len(rows))) if r < len(rows) and i < len(rows[r])]
            col_values = [v for v in col_values if v.strip()]

            # Guess header based on common patterns
            has_prices = any(re.match(r'^\$?[\d,]+\.?\d*$', v) for v in col_values)
            has_qty = any(re.match(r'^\d+$', v) and int(v) < 10000 for v in col_values if re.match(r'^\d+$', v))
            has_codes = any(re.match(r'^[A-Z]{2,}[\w\-]+', v) for v in col_values)

            if has_prices:
                headers.append(f"Price/Amt {i+1}")
            elif has_qty and not has_prices:
                headers.append(f"Qty {i+1}")
            elif has_codes:
                headers.append(f"Code {i+1}")
            else:
                headers.append(f"Col {i+1}")

        return headers

    def _is_data_value(self, text: str) -> bool:
        """Check if text looks like a data value (number, code, price)."""
        text = text.strip()
        # Price
        if re.match(r'^\$?[\d,]+\.?\d*$', text):
            return True
        # Part number / code
        if re.match(r'^[A-Z0-9][\w\-\.]+$', text, re.IGNORECASE) and len(text) > 3:
            return True
        return False

    def on_selection_changed(self):
        """Handle selection change - update UI."""
        items = self.table.selectedItems()
        if items:
            # Get the first selected item
            item = items[0]
            self.selected_value = item.text()
            self.selected_row = item.row()
            self.selected_col = item.column()

            if self.selected_value.strip():
                display = self.selected_value[:40] + "..." if len(self.selected_value) > 40 else self.selected_value
                self.selection_label.setText(f"Selected: \"{display}\" (Row {self.selected_row + 1})")
                self.selection_label.setStyleSheet("color: #2980b9; font-weight: bold;")
                self.btn_add_field.setEnabled(True)
            else:
                self.selection_label.setText("Empty cell selected")
                self.selection_label.setStyleSheet("color: #666; font-style: italic;")
                self.btn_add_field.setEnabled(False)
        else:
            self.selected_value = ""
            self.selected_row = -1
            self.selected_col = -1
            self.selection_label.setText("Click a cell to select, double-click to add as field")
            self.selection_label.setStyleSheet("color: #666; font-style: italic;")
            self.btn_add_field.setEnabled(False)

    def on_cell_double_clicked(self, row: int, col: int):
        """Handle double-click - add as field."""
        item = self.table.item(row, col)
        if item and item.text().strip():
            self.selected_value = item.text()
            self.selected_row = row
            self.selected_col = col
            self.add_selected_as_field()

    def add_selected_as_field(self):
        """Request to add selected cell as a field."""
        if self.selected_value.strip():
            self.field_add_requested.emit(self.selected_value, self.selected_row, self.selected_col)

    def get_selected_text(self) -> str:
        """Get currently selected cell text."""
        return self.selected_value

    def get_row_values(self, row_idx: int) -> List[str]:
        """Get all values from a specific row."""
        if 0 <= row_idx < len(self.rows_data):
            return self.rows_data[row_idx]
        return []

    def get_column_values(self, col_idx: int) -> List[str]:
        """Get all values from a specific column."""
        values = []
        for row in self.rows_data:
            if col_idx < len(row):
                values.append(row[col_idx])
        return values


class ExtractedTextViewer(QTextEdit):
    """Text viewer with selection highlighting for field definition."""

    field_defined = pyqtSignal(str, str, str, str)  # sample_value, context_before, context_after, full_text

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(False)  # Allow selection
        self.setFont(QFont("Consolas", 10))
        self.setLineWrapMode(QTextEdit.WidgetWidth)
        self.defined_fields: List[DefinedField] = []
        self.full_text = ""

    def set_text(self, text: str):
        """Set the extracted text."""
        self.full_text = text
        self.setPlainText(text)

    def get_selection_with_context(self, context_chars: int = 50) -> Tuple[str, str, str]:
        """Get selected text with surrounding context."""
        cursor = self.textCursor()
        if not cursor.hasSelection():
            return "", "", ""

        selected_text = cursor.selectedText()
        start_pos = cursor.selectionStart()
        end_pos = cursor.selectionEnd()

        # Get context before
        context_start = max(0, start_pos - context_chars)
        context_before = self.full_text[context_start:start_pos]
        # Clean up to start at word boundary
        if context_before and context_start > 0:
            first_space = context_before.find(' ')
            if first_space > 0:
                context_before = context_before[first_space + 1:]

        # Get context after
        context_end = min(len(self.full_text), end_pos + context_chars)
        context_after = self.full_text[end_pos:context_end]
        # Clean up to end at word boundary
        if context_after and context_end < len(self.full_text):
            last_space = context_after.rfind(' ')
            if last_space > 0:
                context_after = context_after[:last_space]

        return selected_text, context_before.strip(), context_after.strip()

    def highlight_field(self, field: DefinedField):
        """Highlight a defined field in the text."""
        # Find the sample value in text and highlight it
        fmt = QTextCharFormat()
        fmt.setBackground(field.color)

        # Search and highlight using document find
        doc = self.document()
        cursor = QTextCursor(doc)

        while True:
            cursor = doc.find(field.sample_value, cursor)
            if cursor.isNull():
                break
            cursor.mergeCharFormat(fmt)

    def clear_highlights(self):
        """Clear all highlights by resetting the text."""
        # Save current text and cursor position
        text = self.full_text
        scroll_pos = self.verticalScrollBar().value() if self.verticalScrollBar() else 0

        # Block signals to avoid triggering events
        self.blockSignals(True)
        self.setPlainText(text)
        self.blockSignals(False)

        # Restore scroll position
        if self.verticalScrollBar():
            self.verticalScrollBar().setValue(scroll_pos)

    def refresh_highlights(self, fields: List[DefinedField]):
        """Refresh highlights for all defined fields."""
        if not self.full_text:
            return
        self.clear_highlights()
        for field in fields:
            self.highlight_field(field)


class FieldDefinitionPanel(QWidget):
    """Panel for managing defined fields."""

    field_added = pyqtSignal(object)  # DefinedField
    field_removed = pyqtSignal(int)
    field_updated = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.fields: List[DefinedField] = []
        self.current_index = -1
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)

        # Instructions
        instructions = QLabel(
            "1. Select text in the extracted text area\n"
            "2. Click 'Add Field' to define a field\n"
            "3. Name the field and select its type\n"
            "4. Pattern is auto-generated from context"
        )
        instructions.setStyleSheet("color: #666; font-style: italic; padding: 5px; background: #f8f8f8; border-radius: 3px;")
        instructions.setWordWrap(True)
        layout.addWidget(instructions)

        # Field list
        list_label = QLabel("Defined Fields:")
        list_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(list_label)

        self.field_list = QListWidget()
        self.field_list.currentRowChanged.connect(self.on_field_selected)
        layout.addWidget(self.field_list)

        # Add/Remove buttons
        btn_layout = QHBoxLayout()

        self.btn_add = QPushButton("+ Add Field")
        self.btn_add.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold;")
        self.btn_add.clicked.connect(self.add_field_from_selection)
        btn_layout.addWidget(self.btn_add)

        self.btn_remove = QPushButton("- Remove")
        self.btn_remove.setStyleSheet("background-color: #e74c3c; color: white;")
        self.btn_remove.clicked.connect(self.remove_current_field)
        self.btn_remove.setEnabled(False)
        btn_layout.addWidget(self.btn_remove)

        layout.addLayout(btn_layout)

        # Field properties group
        props_group = QGroupBox("Field Properties")
        props_layout = QFormLayout(props_group)

        self.field_name = QLineEdit()
        self.field_name.setPlaceholderText("Field name")
        self.field_name.textChanged.connect(self.on_name_changed)
        props_layout.addRow("Name:", self.field_name)

        self.field_type = QComboBox()
        self.field_type.addItems([
            "invoice_number",
            "project_number",
            "manufacturer",
            "line_item",
            "indicator",
            "text"
        ])
        self.field_type.currentTextChanged.connect(self.on_type_changed)
        props_layout.addRow("Type:", self.field_type)

        layout.addWidget(props_group)

        # Sample value display
        sample_group = QGroupBox("Captured Value")
        sample_layout = QVBoxLayout(sample_group)

        self.sample_value = QLineEdit()
        self.sample_value.setReadOnly(True)
        self.sample_value.setStyleSheet("background-color: #ecf0f1;")
        sample_layout.addWidget(self.sample_value)

        layout.addWidget(sample_group)

        # Context display
        context_group = QGroupBox("Context (for pattern)")
        context_layout = QVBoxLayout(context_group)

        context_layout.addWidget(QLabel("Before:"))
        self.context_before = QLineEdit()
        self.context_before.setFont(QFont("Consolas", 9))
        self.context_before.setStyleSheet("background-color: #fef9e7;")
        self.context_before.textChanged.connect(self.on_context_changed)
        context_layout.addWidget(self.context_before)

        context_layout.addWidget(QLabel("After:"))
        self.context_after = QLineEdit()
        self.context_after.setFont(QFont("Consolas", 9))
        self.context_after.setStyleSheet("background-color: #fef9e7;")
        self.context_after.textChanged.connect(self.on_context_changed)
        context_layout.addWidget(self.context_after)

        layout.addWidget(context_group)

        # Generated pattern
        pattern_group = QGroupBox("Generated Pattern")
        pattern_layout = QVBoxLayout(pattern_group)

        self.pattern_edit = QLineEdit()
        self.pattern_edit.setFont(QFont("Consolas", 9))
        self.pattern_edit.setPlaceholderText("Auto-generated regex pattern")
        self.pattern_edit.textChanged.connect(self.on_pattern_changed)
        pattern_layout.addWidget(self.pattern_edit)

        # Test result
        self.test_result = QLabel("")
        self.test_result.setWordWrap(True)
        self.test_result.setStyleSheet("color: #3498db; padding: 3px;")
        pattern_layout.addWidget(self.test_result)

        # Regenerate pattern button
        self.btn_regenerate = QPushButton("Regenerate Pattern")
        self.btn_regenerate.clicked.connect(self.regenerate_pattern)
        pattern_layout.addWidget(self.btn_regenerate)

        layout.addWidget(pattern_group)

        layout.addStretch()

        # Store reference to text viewer (set by parent)
        self.text_viewer: Optional[ExtractedTextViewer] = None
        self.full_text = ""

    def set_text_viewer(self, viewer: ExtractedTextViewer):
        """Set reference to text viewer."""
        self.text_viewer = viewer

    def set_full_text(self, text: str):
        """Set the full extracted text for pattern testing."""
        self.full_text = text

    def add_field_from_selection(self):
        """Add a new field from current text selection."""
        if not self.text_viewer:
            return

        sample_value, context_before, context_after = self.text_viewer.get_selection_with_context()

        if not sample_value:
            QMessageBox.warning(
                self, "No Selection",
                "Please select text in the extracted text area first."
            )
            return

        # Ask for field name
        name, ok = QInputDialog.getText(
            self, "Field Name",
            f"Enter a name for this field:\n\nSelected: '{sample_value[:50]}...'" if len(sample_value) > 50 else f"Enter a name for this field:\n\nSelected: '{sample_value}'"
        )

        if not ok or not name:
            return

        # Clean up name
        name = name.strip().lower().replace(' ', '_')
        name = re.sub(r'[^a-z0-9_]', '', name)

        # Check for duplicate names
        if any(f.name == name for f in self.fields):
            QMessageBox.warning(self, "Duplicate Name", f"A field named '{name}' already exists.")
            return

        # Create field
        field = DefinedField(
            name=name,
            field_type=self._guess_field_type(name, sample_value),
            sample_value=sample_value,
            context_before=context_before,
            context_after=context_after
        )

        # Generate initial pattern
        field.pattern = self._generate_pattern(field)

        self.fields.append(field)
        self.refresh_list()

        # Select the new field
        self.field_list.setCurrentRow(len(self.fields) - 1)

        # Highlight in text viewer
        if self.text_viewer:
            self.text_viewer.refresh_highlights(self.fields)

        self.field_added.emit(field)

    def _guess_field_type(self, name: str, value: str) -> str:
        """Guess field type based on name and value."""
        name_lower = name.lower()
        value_lower = value.lower()

        if 'invoice' in name_lower or 'inv' in name_lower:
            return 'invoice_number'
        elif 'project' in name_lower or 'po' in name_lower or 'order' in name_lower:
            return 'project_number'
        elif 'manufacturer' in name_lower or 'supplier' in name_lower or 'vendor' in name_lower:
            return 'manufacturer'
        elif 'item' in name_lower or 'line' in name_lower or 'part' in name_lower:
            return 'line_item'
        elif 'indicator' in name_lower or 'id' in name_lower:
            return 'indicator'

        # Check value patterns
        if re.match(r'^[A-Z]{2,3}[-/]?\d+', value):
            return 'invoice_number'
        elif re.match(r'^\d{6,}$', value):
            return 'project_number'

        return 'text'

    def _generate_pattern(self, field: DefinedField) -> str:
        """Generate regex pattern from field context."""
        # Clean and prepare context - extract only useful nearby words
        before_clean = self._clean_context(field.context_before) if field.context_before else ""
        after_clean = self._clean_context(field.context_after) if field.context_after else ""

        # Generate capture group for value
        capture = self._generalize_value(field.sample_value)

        # Build pattern - use only a few words of context
        parts = []
        if before_clean:
            # Get last few words for context matching
            words = before_clean.split()
            if words:
                # Use last 2-3 words as context
                context_words = words[-3:] if len(words) >= 3 else words
                before_pattern = r'\s+'.join(self._escape_regex(w) for w in context_words)
                parts.append(before_pattern)
                parts.append(r'\s+')

        parts.append(capture)

        if after_clean:
            # Get first few words for context matching
            words = after_clean.split()
            if words:
                context_words = words[:3] if len(words) >= 3 else words
                after_pattern = r'\s+'.join(self._escape_regex(w) for w in context_words)
                parts.append(r'\s+')
                parts.append(after_pattern)

        return ''.join(parts)

    def _clean_context(self, text: str) -> str:
        """Clean context text for use in pattern - remove newlines, extra spaces."""
        if not text:
            return ""
        # Replace newlines and multiple spaces with single space
        cleaned = re.sub(r'[\n\r\t]+', ' ', text)
        cleaned = re.sub(r'\s+', ' ', cleaned)
        # Remove non-alphanumeric except common punctuation
        cleaned = re.sub(r'[^\w\s\.\-:/#]', '', cleaned)
        return cleaned.strip()

    def _escape_regex(self, text: str) -> str:
        """Escape special regex characters."""
        special = r'\.^$*+?{}[]|()'
        result = ""
        for c in text:
            if c in special:
                result += '\\' + c
            else:
                result += c
        return result

    def _generalize_value(self, value: str) -> str:
        """Convert literal value to regex capture group."""
        value = value.strip()
        if not value:
            return r'(.+?)'

        # Alphanumeric with dashes (invoice/reference numbers)
        if re.match(r'^[A-Z0-9][\w\-/\.]+$', value, re.IGNORECASE):
            return r'([A-Za-z0-9][\w\-/\.]+)'

        # Pure numbers with possible commas/decimals
        if re.match(r'^[\d,]+\.?\d*$', value):
            return r'([\d,]+\.?\d*)'

        # Currency
        if re.match(r'^\$?[\d,]+\.?\d*$', value):
            return r'\$?([\d,]+\.?\d*)'

        # Date formats
        if re.match(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$', value):
            return r'(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})'
        if re.match(r'^\d{4}[/\-]\d{2}[/\-]\d{2}$', value):
            return r'(\d{4}[/\-]\d{2}[/\-]\d{2})'

        # HTS codes
        if re.match(r'^\d{4}\.\d{2}\.\d{4}$', value):
            return r'(\d{4}\.\d{2}\.\d{4})'

        # General text with words
        if ' ' in value:
            return r'(.+?)'

        # Single word
        return r'(\S+)'

    def refresh_list(self):
        """Refresh the field list."""
        self.field_list.clear()
        for field in self.fields:
            item_text = f"{field.name} ({field.field_type})"
            item = QListWidgetItem(item_text)
            item.setBackground(field.color)
            self.field_list.addItem(item)

        self.btn_remove.setEnabled(len(self.fields) > 0)

    def on_field_selected(self, index: int):
        """Handle field selection."""
        self.current_index = index
        if 0 <= index < len(self.fields):
            field = self.fields[index]

            # Block signals while updating
            self.field_name.blockSignals(True)
            self.field_type.blockSignals(True)
            self.pattern_edit.blockSignals(True)
            self.context_before.blockSignals(True)
            self.context_after.blockSignals(True)

            self.field_name.setText(field.name)
            self.field_type.setCurrentText(field.field_type)
            self.sample_value.setText(field.sample_value)
            self.context_before.setText(field.context_before)
            self.context_after.setText(field.context_after)
            self.pattern_edit.setText(field.pattern)

            self.field_name.blockSignals(False)
            self.field_type.blockSignals(False)
            self.pattern_edit.blockSignals(False)
            self.context_before.blockSignals(False)
            self.context_after.blockSignals(False)

            # Test the pattern
            self.test_pattern()

            self.btn_remove.setEnabled(True)
        else:
            self.btn_remove.setEnabled(False)

    def on_name_changed(self, text: str):
        """Handle field name change."""
        if 0 <= self.current_index < len(self.fields):
            self.fields[self.current_index].name = text
            self.refresh_list()
            self.field_list.setCurrentRow(self.current_index)
            self.field_updated.emit()

    def on_type_changed(self, text: str):
        """Handle field type change."""
        if 0 <= self.current_index < len(self.fields):
            self.fields[self.current_index].field_type = text
            self.fields[self.current_index].color = self.fields[self.current_index]._get_color_for_type(text)
            self.refresh_list()
            self.field_list.setCurrentRow(self.current_index)
            if self.text_viewer:
                self.text_viewer.refresh_highlights(self.fields)
            self.field_updated.emit()

    def on_context_changed(self):
        """Handle context change."""
        if 0 <= self.current_index < len(self.fields):
            self.fields[self.current_index].context_before = self.context_before.text()
            self.fields[self.current_index].context_after = self.context_after.text()

    def on_pattern_changed(self, text: str):
        """Handle pattern change."""
        if 0 <= self.current_index < len(self.fields):
            self.fields[self.current_index].pattern = text
            self.test_pattern()

    def regenerate_pattern(self):
        """Regenerate pattern from current context."""
        if 0 <= self.current_index < len(self.fields):
            field = self.fields[self.current_index]
            field.pattern = self._generate_pattern(field)
            self.pattern_edit.setText(field.pattern)

    def test_pattern(self):
        """Test current pattern against full text."""
        if not self.full_text:
            self.test_result.setText("No text loaded")
            self.test_result.setStyleSheet("color: orange;")
            return

        pattern = self.pattern_edit.text()
        if not pattern:
            self.test_result.setText("No pattern")
            self.test_result.setStyleSheet("color: orange;")
            return

        try:
            matches = re.findall(pattern, self.full_text, re.IGNORECASE | re.MULTILINE)
            if matches:
                # Show first match
                first = matches[0]
                if isinstance(first, tuple):
                    first = first[0] if first else ""
                display = first[:40] + "..." if len(str(first)) > 40 else str(first)
                self.test_result.setText(f"✓ {len(matches)} match(es): '{display}'")
                self.test_result.setStyleSheet("color: green;")
            else:
                self.test_result.setText("✗ No matches found")
                self.test_result.setStyleSheet("color: red;")
        except re.error as e:
            self.test_result.setText(f"Regex error: {e}")
            self.test_result.setStyleSheet("color: red;")

    def remove_current_field(self):
        """Remove the current field."""
        if 0 <= self.current_index < len(self.fields):
            del self.fields[self.current_index]
            self.refresh_list()
            if self.text_viewer:
                self.text_viewer.refresh_highlights(self.fields)
            self.field_removed.emit(self.current_index)
            self.current_index = -1


class DetectedPatternsDialog(QDialog):
    """Dialog to display and select detected patterns."""

    def __init__(self, patterns: List[DetectedPattern], full_text: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Detected Patterns")
        self.setMinimumSize(800, 600)
        self.patterns = patterns
        self.full_text = full_text
        self.checkboxes = []
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Header
        header = QLabel("Detected Patterns")
        header.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(header)

        desc = QLabel(
            "The following patterns were detected in your invoice.\n"
            "Select the patterns you want to use as field definitions."
        )
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # Scrollable pattern list
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout(content)

        # Group by pattern type
        line_items = [p for p in self.patterns if p.pattern_type == 'line_item']
        invoice_nums = [p for p in self.patterns if p.pattern_type == 'invoice_number']
        po_nums = [p for p in self.patterns if p.pattern_type == 'project_number']

        if line_items:
            group = QGroupBox("Line Item Patterns")
            group.setStyleSheet("QGroupBox { font-weight: bold; color: #e67e22; }")
            group_layout = QVBoxLayout(group)
            for pattern in line_items:
                self._add_pattern_widget(group_layout, pattern)
            content_layout.addWidget(group)

        if invoice_nums:
            group = QGroupBox("Invoice Number Patterns")
            group.setStyleSheet("QGroupBox { font-weight: bold; color: #2ecc71; }")
            group_layout = QVBoxLayout(group)
            for pattern in invoice_nums:
                self._add_pattern_widget(group_layout, pattern)
            content_layout.addWidget(group)

        if po_nums:
            group = QGroupBox("PO/Project Number Patterns")
            group.setStyleSheet("QGroupBox { font-weight: bold; color: #9b59b6; }")
            group_layout = QVBoxLayout(group)
            for pattern in po_nums:
                self._add_pattern_widget(group_layout, pattern)
            content_layout.addWidget(group)

        content_layout.addStretch()
        scroll.setWidget(content)
        layout.addWidget(scroll, 1)

        # Buttons
        btn_layout = QHBoxLayout()

        btn_select_all = QPushButton("Select All")
        btn_select_all.clicked.connect(self._select_all)
        btn_layout.addWidget(btn_select_all)

        btn_select_none = QPushButton("Select None")
        btn_select_none.clicked.connect(self._select_none)
        btn_layout.addWidget(btn_select_none)

        btn_layout.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_add = QPushButton("Add Selected")
        btn_add.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px 16px;")
        btn_add.clicked.connect(self.accept)
        btn_layout.addWidget(btn_add)

        layout.addLayout(btn_layout)

    def _add_pattern_widget(self, layout: QVBoxLayout, pattern: DetectedPattern):
        """Add a pattern with checkbox to the layout."""
        widget = QWidget()
        widget_layout = QVBoxLayout(widget)
        widget_layout.setContentsMargins(10, 5, 10, 5)

        # Checkbox with description
        checkbox = QCheckBox(pattern.description)
        checkbox.setFont(QFont("Arial", 10, QFont.Bold))
        checkbox.setChecked(pattern.confidence > 0.5)  # Auto-select high confidence
        self.checkboxes.append((checkbox, pattern))
        widget_layout.addWidget(checkbox)

        # Confidence indicator
        confidence_text = f"Confidence: {pattern.confidence:.0%}"
        if pattern.confidence >= 0.7:
            confidence_style = "color: #27ae60;"  # Green
        elif pattern.confidence >= 0.4:
            confidence_style = "color: #f39c12;"  # Orange
        else:
            confidence_style = "color: #e74c3c;"  # Red
        confidence_label = QLabel(confidence_text)
        confidence_label.setStyleSheet(confidence_style)
        widget_layout.addWidget(confidence_label)

        # Sample matches
        if pattern.sample_matches:
            samples_label = QLabel("Sample matches:")
            samples_label.setStyleSheet("color: #666; font-style: italic;")
            widget_layout.addWidget(samples_label)

            samples_text = QPlainTextEdit()
            samples_text.setFont(QFont("Consolas", 9))
            samples_text.setPlainText('\n'.join(pattern.sample_matches[:3]))
            samples_text.setMaximumHeight(70)
            samples_text.setReadOnly(True)
            widget_layout.addWidget(samples_text)

        # Pattern regex
        pattern_label = QLabel(f"Pattern: {pattern.pattern[:80]}...")
        pattern_label.setStyleSheet("color: #3498db; font-family: Consolas; font-size: 9px;")
        pattern_label.setWordWrap(True)
        widget_layout.addWidget(pattern_label)

        widget.setStyleSheet("QWidget { background-color: #f8f9fa; border-radius: 5px; margin: 5px; }")
        layout.addWidget(widget)

    def _select_all(self):
        """Select all checkboxes."""
        for checkbox, _ in self.checkboxes:
            checkbox.setChecked(True)

    def _select_none(self):
        """Deselect all checkboxes."""
        for checkbox, _ in self.checkboxes:
            checkbox.setChecked(False)

    def get_selected_patterns(self) -> List[Dict]:
        """Get the selected patterns as field definitions."""
        selected = []
        for checkbox, pattern in self.checkboxes:
            if checkbox.isChecked():
                # Generate field name from pattern type
                base_name = pattern.pattern_type
                count = sum(1 for s in selected if s['type'] == pattern.pattern_type)
                name = f"{base_name}_{count + 1}" if count > 0 else base_name

                # Get sample value
                sample = pattern.sample_matches[0] if pattern.sample_matches else ""

                selected.append({
                    'name': name,
                    'type': pattern.pattern_type,
                    'pattern': pattern.pattern,
                    'sample': sample
                })
        return selected


class VisualTemplateBuilderDialog(QDialog):
    """
    Visual Template Builder with text-based field definition.
    Load PDF → Extract text → Highlight text to define fields → Generate template.
    """

    template_created = pyqtSignal(str, str)  # template_name, file_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Visual Template Builder")
        self.setWindowFlags(
            Qt.Window |
            Qt.WindowMinimizeButtonHint |
            Qt.WindowMaximizeButtonHint |
            Qt.WindowCloseButtonHint
        )
        self.setMinimumSize(1200, 800)
        self.resize(1400, 900)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.pdf_path = ""
        self.full_text = ""

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Header
        header_layout = QHBoxLayout()

        title = QLabel("Visual Template Builder")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        header_layout.addWidget(title)

        header_layout.addStretch()

        # PDF selection
        self.pdf_label = QLabel("No PDF loaded")
        self.pdf_label.setStyleSheet("color: #666;")
        header_layout.addWidget(self.pdf_label)

        self.btn_open = QPushButton("Open PDF")
        self.btn_open.setStyleSheet("font-weight: bold; padding: 8px 16px;")
        self.btn_open.clicked.connect(self.open_pdf)
        header_layout.addWidget(self.btn_open)

        self.btn_detect = QPushButton("Auto-Detect Patterns")
        self.btn_detect.setStyleSheet("font-weight: bold; padding: 8px 16px; background-color: #9b59b6; color: white;")
        self.btn_detect.clicked.connect(self.auto_detect_patterns)
        self.btn_detect.setEnabled(False)
        self.btn_detect.setToolTip("Automatically detect line item patterns, invoice numbers, and PO numbers")
        header_layout.addWidget(self.btn_detect)

        layout.addLayout(header_layout)

        # Main content splitter
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        # Left side: Extracted text with view toggle
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Header with view toggle
        header_row = QHBoxLayout()
        text_header = QLabel("Extracted Text (select text to define fields)")
        text_header.setStyleSheet("font-weight: bold; font-size: 12px;")
        header_row.addWidget(text_header)

        header_row.addStretch()

        # View toggle buttons
        self.view_toggle_group = QButtonGroup(self)
        self.btn_raw_view = QRadioButton("Raw Text")
        self.btn_raw_view.setChecked(True)
        self.btn_raw_view.setStyleSheet("font-size: 11px;")
        self.view_toggle_group.addButton(self.btn_raw_view, 0)
        header_row.addWidget(self.btn_raw_view)

        self.btn_table_view = QRadioButton("Table View")
        self.btn_table_view.setStyleSheet("font-size: 11px;")
        self.view_toggle_group.addButton(self.btn_table_view, 1)
        header_row.addWidget(self.btn_table_view)

        self.view_toggle_group.buttonClicked.connect(self.toggle_text_view)

        left_layout.addLayout(header_row)

        # Stacked widget for view switching
        self.view_stack = QWidget()
        stack_layout = QVBoxLayout(self.view_stack)
        stack_layout.setContentsMargins(0, 0, 0, 0)

        # Raw text viewer
        self.text_viewer = ExtractedTextViewer()
        self.text_viewer.setPlaceholderText(
            "Load a PDF to extract text...\n\n"
            "Once loaded:\n"
            "1. Select (highlight) text that represents a field value\n"
            "2. Click '+ Add Field' on the right panel\n"
            "3. Name the field and select its type\n"
            "4. The pattern will be auto-generated\n\n"
            "TIP: Use 'Table View' to see the document structure with columns"
        )
        stack_layout.addWidget(self.text_viewer)

        # Table viewer (initially hidden)
        self.table_viewer = TableTextViewer()
        self.table_viewer.hide()
        self.table_viewer.field_add_requested.connect(self.on_table_field_requested)
        stack_layout.addWidget(self.table_viewer)

        left_layout.addWidget(self.view_stack)

        splitter.addWidget(left_widget)

        # Right side: Field definition panel
        right_widget = QWidget()
        right_widget.setMinimumWidth(400)
        right_widget.setMaximumWidth(500)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.field_panel = FieldDefinitionPanel()
        self.field_panel.set_text_viewer(self.text_viewer)
        self.field_panel.field_added.connect(self.on_field_added)
        self.field_panel.field_removed.connect(self.on_field_removed)
        self.field_panel.field_updated.connect(self.on_field_updated)
        right_layout.addWidget(self.field_panel, 1)

        # Template settings
        settings_group = QGroupBox("Template Settings")
        settings_layout = QFormLayout(settings_group)

        self.template_name = QLineEdit()
        self.template_name.setPlaceholderText("e.g., acme_corp")
        settings_layout.addRow("Template Name:", self.template_name)

        self.company_name = QLineEdit()
        self.company_name.setPlaceholderText("e.g., Acme Corporation")
        settings_layout.addRow("Company Name:", self.company_name)

        self.template_desc = QLineEdit()
        self.template_desc.setPlaceholderText("Brief description")
        settings_layout.addRow("Description:", self.template_desc)

        right_layout.addWidget(settings_group)

        # Generate button
        self.btn_generate = QPushButton("Generate Template")
        self.btn_generate.setStyleSheet(
            "font-weight: bold; padding: 15px; font-size: 14px; "
            "background-color: #3498db; color: white;"
        )
        self.btn_generate.clicked.connect(self.generate_template)
        self.btn_generate.setEnabled(False)
        right_layout.addWidget(self.btn_generate)

        splitter.addWidget(right_widget)

        # Set splitter sizes
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([900, 400])

        layout.addWidget(splitter, 1)

        # Status bar
        self.status_bar = QStatusBar()
        layout.addWidget(self.status_bar)
        self.status_bar.showMessage("Open a PDF to begin")

    def open_pdf(self):
        """Open a PDF file and extract text."""
        if not HAS_PDFPLUMBER:
            QMessageBox.warning(
                self, "Missing Dependency",
                "pdfplumber is not installed.\n\nRun: pip install pdfplumber"
            )
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open PDF Invoice",
            "", "PDF Files (*.pdf)"
        )

        if file_path:
            self.load_pdf(file_path)

    def load_pdf(self, file_path: str):
        """Load PDF and extract all text."""
        self.pdf_path = file_path
        self.status_bar.showMessage("Loading PDF...")

        try:
            # Extract text with pdfplumber
            with pdfplumber.open(file_path) as pdf:
                texts = []
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    texts.append(f"--- Page {i+1} ---\n{text}")

                self.full_text = "\n\n".join(texts)

            # Display text in both viewers
            self.text_viewer.set_text(self.full_text)
            self.table_viewer.set_text(self.full_text)
            self.field_panel.set_full_text(self.full_text)

            # Update UI
            filename = Path(file_path).name
            self.pdf_label.setText(f"Loaded: {filename}")
            self.pdf_label.setStyleSheet("color: #27ae60; font-weight: bold;")

            # Auto-detect company name
            self._auto_detect_company()

            # Enable buttons
            self.btn_generate.setEnabled(True)
            self.btn_detect.setEnabled(True)

            self.status_bar.showMessage(f"Loaded: {filename} - {len(self.full_text)} characters extracted")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load PDF: {e}")
            self.status_bar.showMessage("Error loading PDF")

    def _auto_detect_company(self):
        """Try to auto-detect company name from text."""
        patterns = [
            r'(?:From|Supplier|Seller|Vendor|Shipper)[\s:]+([A-Z][A-Za-z\s&\.]+(?:Ltd|LLC|Inc|Corp|Co\.|GmbH|s\.r\.o\.|Limited|CORPORATION))',
            r'^([A-Z][A-Z\s&\.]+(?:LTD|LLC|INC|CORP|CO\.|GMBH|LIMITED|CORPORATION))\s*$',
        ]

        for pattern in patterns:
            match = re.search(pattern, self.full_text, re.MULTILINE | re.IGNORECASE)
            if match:
                company = match.group(1).strip()
                if 3 < len(company) < 50:
                    self.company_name.setText(company)
                    # Generate template name
                    name = company.lower()
                    name = re.sub(r'[^a-z0-9]+', '_', name)
                    name = re.sub(r'_+', '_', name).strip('_')[:30]
                    self.template_name.setText(name)
                    return

    def on_field_added(self, field: DefinedField):
        """Handle new field added."""
        self.status_bar.showMessage(f"Field '{field.name}' added ({field.field_type})")

    def on_field_removed(self, index: int):
        """Handle field removed."""
        self.status_bar.showMessage("Field removed")

    def on_field_updated(self):
        """Handle field updated."""
        pass

    def toggle_text_view(self, button):
        """Toggle between raw text and table view."""
        if button == self.btn_raw_view:
            self.text_viewer.show()
            self.table_viewer.hide()
            self.status_bar.showMessage("Switched to Raw Text view")
        else:
            self.text_viewer.hide()
            self.table_viewer.show()
            if self.full_text:
                self.table_viewer.set_text(self.full_text)
            self.status_bar.showMessage("Switched to Table View - click cells to select values")

    def on_table_field_requested(self, value: str, row: int, col: int):
        """Handle request to add field from table view."""
        if not value.strip():
            return

        # Get row context for better pattern generation
        row_values = self.table_viewer.get_row_values(row)
        col_values = self.table_viewer.get_column_values(col)

        # Show add field dialog with more context
        display_val = value[:50] + "..." if len(value) > 50 else value
        row_context = " | ".join(v for v in row_values if v.strip())[:80]

        name, ok = QInputDialog.getText(
            self, "Add Field from Table",
            f"Selected value: \"{display_val}\"\n"
            f"Row context: {row_context}...\n\n"
            f"Enter a name for this field:"
        )

        if not ok or not name:
            return

        # Clean up name
        name = name.strip().lower().replace(' ', '_')
        name = re.sub(r'[^a-z0-9_]', '', name)

        # Check for duplicate names
        if any(f.name == name for f in self.field_panel.fields):
            QMessageBox.warning(self, "Duplicate Name", f"A field named '{name}' already exists.")
            return

        # Create field with column context
        field = DefinedField(
            name=name,
            field_type=self.field_panel._guess_field_type(name, value),
            sample_value=value,
            context_before=f"Column {col+1}",
            context_after=""
        )

        # Generate pattern based on the value type
        field.pattern = self.field_panel._generalize_value(value)

        self.field_panel.fields.append(field)
        self.field_panel.refresh_list()

        # Select the new field
        self.field_panel.field_list.setCurrentRow(len(self.field_panel.fields) - 1)

        # Update highlights in raw view
        self.text_viewer.refresh_highlights(self.field_panel.fields)

        self.field_panel.field_added.emit(field)
        self.status_bar.showMessage(f"Field '{name}' added from table (Row {row+1}, Col {col+1})")

    def auto_detect_patterns(self):
        """Auto-detect patterns in the extracted text."""
        if not self.full_text:
            QMessageBox.warning(self, "No Text", "Please load a PDF first.")
            return

        self.status_bar.showMessage("Detecting patterns...")

        # Detect patterns
        detected = PatternDetector.detect_patterns(self.full_text)

        if not detected:
            QMessageBox.information(
                self, "No Patterns Found",
                "No common patterns were detected in the text.\n\n"
                "You can still manually define fields by selecting text."
            )
            self.status_bar.showMessage("No patterns detected")
            return

        # Show detected patterns dialog
        dialog = DetectedPatternsDialog(detected, self.full_text, self)
        if dialog.exec_() == QDialog.Accepted:
            # Get selected patterns and add as fields
            selected = dialog.get_selected_patterns()
            for pattern_info in selected:
                field = DefinedField(
                    name=pattern_info['name'],
                    field_type=pattern_info['type'],
                    sample_value=pattern_info['sample'],
                    context_before="",
                    context_after=""
                )
                field.pattern = pattern_info['pattern']
                self.field_panel.fields.append(field)

            self.field_panel.refresh_list()
            self.text_viewer.refresh_highlights(self.field_panel.fields)
            self.status_bar.showMessage(f"Added {len(selected)} pattern(s) as fields")
        else:
            self.status_bar.showMessage("Pattern detection cancelled")

    def generate_template(self):
        """Generate the template code from defined fields."""
        template_name = self.template_name.text().strip().lower()
        template_name = re.sub(r'[^a-z0-9_]', '', template_name)

        if not template_name:
            QMessageBox.warning(self, "Missing Name", "Please enter a template name.")
            return

        fields = self.field_panel.fields
        if not fields:
            QMessageBox.warning(
                self, "No Fields",
                "Please define at least one field by selecting text."
            )
            return

        company = self.company_name.text().strip() or "Unknown Company"
        description = self.template_desc.text().strip() or f"Template for {company}"

        # Categorize fields
        invoice_field = None
        project_field = None
        manufacturer_field = None
        indicators = []
        line_item_fields = []
        other_fields = []

        for field in fields:
            if field.field_type == "invoice_number":
                invoice_field = field
            elif field.field_type == "project_number":
                project_field = field
            elif field.field_type == "manufacturer":
                manufacturer_field = field
            elif field.field_type == "indicator":
                indicators.append(field)
            elif field.field_type == "line_item":
                line_item_fields.append(field)
            else:
                other_fields.append(field)

        # Generate class name
        class_name = ''.join(word.capitalize() for word in template_name.split('_')) + 'Template'

        # Generate code
        code = self._generate_template_code(
            template_name, class_name, company, description,
            invoice_field, project_field, manufacturer_field,
            indicators, line_item_fields
        )

        # Show preview and save
        self._preview_and_save(template_name, class_name, code)

    def _generate_template_code(
        self,
        template_name: str,
        class_name: str,
        company: str,
        description: str,
        invoice_field: Optional[DefinedField],
        project_field: Optional[DefinedField],
        manufacturer_field: Optional[DefinedField],
        indicators: List[DefinedField],
        line_item_fields: List[DefinedField]
    ) -> str:
        """Generate the Python template code."""

        # Build indicator checks
        indicator_checks = []
        for ind in indicators:
            if ind.sample_value:
                snippet = ind.sample_value.strip()[:30].lower()
                snippet = re.sub(r"'", "\\'", snippet)
                indicator_checks.append(f"'{snippet}' in text.lower()")

        if not indicator_checks:
            # Use company name as indicator
            company_lower = company.lower().replace("'", "\\'")
            indicator_checks.append(f"'{company_lower}' in text.lower()")

        indicators_code = ' or '.join(indicator_checks)

        # Invoice pattern - patterns should NOT be double-escaped
        # The patterns already have proper escaping for raw strings
        if invoice_field and invoice_field.pattern:
            inv_pattern = invoice_field.pattern.replace("'", "\\'")
        else:
            inv_pattern = r'[Ii]nvoice\s*(?:#|[Nn]o\.?|[Nn]umber)?\s*:?\s*([A-Z0-9][\w\-/]+)'

        # Project pattern
        if project_field and project_field.pattern:
            proj_pattern = project_field.pattern.replace("'", "\\'")
        else:
            proj_pattern = r'(?:[Pp]roject|P\.?O\.?|[Oo]rder)\s*#?\s*:?\s*([A-Z0-9][\w\-]+)'

        # Manufacturer extraction
        if manufacturer_field and manufacturer_field.pattern:
            mfg_pattern = manufacturer_field.pattern.replace("'", "\\'")
            mfg_code = f'''pattern = r'{mfg_pattern}'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip() if match.lastindex else match.group(0).strip()
        return "{company}"'''
        else:
            mfg_code = f'return "{company}"'

        # Line item pattern
        if line_item_fields and line_item_fields[0].pattern:
            line_pattern = line_item_fields[0].pattern.replace("'", "\\'")
        else:
            line_pattern = r'^([A-Z0-9][\w\-\.]+)\s+(\d+(?:[.,]\d+)?)\s+\$?([\d,]+\.?\d*)'

        code = f'''"""
{class_name} - Invoice template for {company}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""

import re
from typing import List, Dict
from .base_template import BaseTemplate


class {class_name}(BaseTemplate):
    """
    Invoice template for {company}.
    Generated by Visual Template Builder.
    """

    name = "{template_name.replace('_', ' ').title()}"
    description = "{description}"
    client = "{company}"
    version = "1.0.0"
    enabled = True

    extra_columns = ['unit_price', 'description']

    def can_process(self, text: str) -> bool:
        """Check if this template can process the invoice."""
        return {indicators_code}

    def get_confidence_score(self, text: str) -> float:
        """Return confidence score for template matching."""
        if not self.can_process(text):
            return 0.0

        score = 0.5
        text_lower = text.lower()

        # Add confidence based on multiple indicators
        if '{company.lower()}' in text_lower:
            score += 0.3
        if re.search(r'{inv_pattern}', text, re.IGNORECASE):
            score += 0.1
        if re.search(r'{proj_pattern}', text, re.IGNORECASE):
            score += 0.1

        return min(score, 1.0)

    def extract_invoice_number(self, text: str) -> str:
        """Extract invoice number."""
        pattern = r'{inv_pattern}'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip() if match.lastindex else match.group(0).strip()
        return "UNKNOWN"

    def extract_project_number(self, text: str) -> str:
        """Extract project number."""
        pattern = r'{proj_pattern}'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip() if match.lastindex else match.group(0).strip()
        return "UNKNOWN"

    def extract_manufacturer_name(self, text: str) -> str:
        """Extract manufacturer name."""
        {mfg_code}

    def extract_line_items(self, text: str) -> List[Dict]:
        """Extract line items from invoice."""
        line_items = []
        seen_items = set()

        # Line item pattern
        pattern = re.compile(r'{line_pattern}', re.MULTILINE | re.IGNORECASE)

        for match in pattern.finditer(text):
            try:
                groups = match.groups()
                part_number = groups[0] if len(groups) > 0 else ''
                quantity = groups[1] if len(groups) > 1 else '1'
                total_price = groups[-1] if len(groups) > 2 else groups[1] if len(groups) > 1 else '0'

                # Clean numeric values
                quantity = re.sub(r'[^\\d.]', '', str(quantity))
                total_price = re.sub(r'[^\\d.]', '', str(total_price))

                # Deduplication
                item_key = f"{{part_number}}_{{quantity}}_{{total_price}}"
                if item_key not in seen_items and part_number:
                    seen_items.add(item_key)

                    qty = float(quantity) if quantity else 1
                    price = float(total_price) if total_price else 0
                    unit_price = price / qty if qty > 0 else price

                    line_items.append({{
                        'part_number': part_number,
                        'quantity': quantity,
                        'total_price': total_price,
                        'unit_price': f"{{unit_price:.2f}}",
                    }})

            except (ValueError, IndexError):
                continue

        return line_items

    def is_packing_list(self, text: str) -> bool:
        """Check if document is a packing list."""
        text_lower = text.lower()
        if 'packing list' in text_lower or 'packing slip' in text_lower:
            if 'invoice' not in text_lower:
                return True
        return False
'''
        return code

    def _validate_patterns(self, code: str) -> List[str]:
        """Validate all regex patterns in the generated code."""
        errors = []
        # Find all raw string patterns: r'...' or r"..."
        pattern_matches = re.findall(r"r['\"]([^'\"]+)['\"]", code)

        for pattern in pattern_matches:
            try:
                re.compile(pattern)
            except re.error as e:
                errors.append(f"Invalid regex: {pattern[:50]}... - {e}")

        return errors

    def _test_template_against_pdf(self, code: str) -> dict:
        """Test the generated template against the source PDF text."""
        results = {
            'can_process': False,
            'confidence': 0.0,
            'invoice_number': 'N/A',
            'project_number': 'N/A',
            'manufacturer': 'N/A',
            'line_items': [],
            'errors': []
        }

        if not self.full_text:
            results['errors'].append("No PDF text available for testing")
            return results

        try:
            # Create a temporary module to test the template
            import types
            import sys

            # Create module
            temp_module = types.ModuleType('temp_template')
            temp_module.__dict__['re'] = re
            temp_module.__dict__['List'] = List
            temp_module.__dict__['Dict'] = Dict

            # We need BaseTemplate - import it
            from templates.base_template import BaseTemplate
            temp_module.__dict__['BaseTemplate'] = BaseTemplate

            # Execute the template code in the module
            exec(code, temp_module.__dict__)

            # Find the template class
            template_class = None
            for name, obj in temp_module.__dict__.items():
                if isinstance(obj, type) and issubclass(obj, BaseTemplate) and obj is not BaseTemplate:
                    template_class = obj
                    break

            if not template_class:
                results['errors'].append("Could not find template class in generated code")
                return results

            # Create instance and test
            template = template_class()

            results['can_process'] = template.can_process(self.full_text)
            results['confidence'] = template.get_confidence_score(self.full_text)
            results['invoice_number'] = template.extract_invoice_number(self.full_text)
            results['project_number'] = template.extract_project_number(self.full_text)
            results['manufacturer'] = template.extract_manufacturer_name(self.full_text)

            # Extract line items
            try:
                items = template.extract_line_items(self.full_text)
                results['line_items'] = items[:10]  # Limit to first 10 for preview
            except Exception as e:
                results['errors'].append(f"Line item extraction failed: {e}")

        except SyntaxError as e:
            results['errors'].append(f"Syntax error in template: {e}")
        except Exception as e:
            results['errors'].append(f"Template test failed: {e}")

        return results

    def _preview_and_save(self, template_name: str, class_name: str, code: str):
        """Show preview dialog with validation and testing."""
        # First validate patterns
        validation_errors = self._validate_patterns(code)

        # Test against PDF
        test_results = self._test_template_against_pdf(code)

        # Create preview dialog
        preview_dialog = QDialog(self)
        preview_dialog.setWindowTitle("Generated Template Preview")
        preview_dialog.setMinimumSize(1000, 700)

        layout = QVBoxLayout(preview_dialog)

        # Create tabs for different views
        tabs = QTabWidget()

        # Tab 1: Test Results
        test_widget = QWidget()
        test_layout = QVBoxLayout(test_widget)

        # Validation status
        if validation_errors:
            val_label = QLabel("⚠️ Pattern Validation Errors:")
            val_label.setStyleSheet("color: #e74c3c; font-weight: bold;")
            test_layout.addWidget(val_label)
            for err in validation_errors:
                err_label = QLabel(f"  • {err}")
                err_label.setStyleSheet("color: #e74c3c;")
                test_layout.addWidget(err_label)
        else:
            val_label = QLabel("✓ All patterns validated successfully")
            val_label.setStyleSheet("color: #27ae60; font-weight: bold;")
            test_layout.addWidget(val_label)

        test_layout.addWidget(QLabel(""))  # Spacer

        # Test results summary
        summary_group = QGroupBox("Template Test Results (Against Source PDF)")
        summary_layout = QFormLayout(summary_group)

        # Can process
        can_process_label = QLabel("✓ Yes" if test_results['can_process'] else "✗ No")
        can_process_label.setStyleSheet(f"color: {'#27ae60' if test_results['can_process'] else '#e74c3c'}; font-weight: bold;")
        summary_layout.addRow("Can Process:", can_process_label)

        # Confidence
        confidence = test_results['confidence']
        conf_color = '#27ae60' if confidence >= 0.7 else '#f39c12' if confidence >= 0.4 else '#e74c3c'
        conf_label = QLabel(f"{confidence:.0%}")
        conf_label.setStyleSheet(f"color: {conf_color}; font-weight: bold;")
        summary_layout.addRow("Confidence Score:", conf_label)

        # Extracted values
        summary_layout.addRow("Invoice Number:", QLabel(str(test_results['invoice_number'])))
        summary_layout.addRow("Project Number:", QLabel(str(test_results['project_number'])))
        summary_layout.addRow("Manufacturer:", QLabel(str(test_results['manufacturer'])))

        test_layout.addWidget(summary_group)

        # Line items preview
        items_group = QGroupBox(f"Extracted Line Items ({len(test_results['line_items'])} found)")
        items_layout = QVBoxLayout(items_group)

        if test_results['line_items']:
            items_table = QTableWidget()
            items_table.setColumnCount(4)
            items_table.setHorizontalHeaderLabels(['Part Number', 'Quantity', 'Total Price', 'Unit Price'])
            items_table.setRowCount(len(test_results['line_items']))

            for i, item in enumerate(test_results['line_items']):
                items_table.setItem(i, 0, QTableWidgetItem(str(item.get('part_number', ''))))
                items_table.setItem(i, 1, QTableWidgetItem(str(item.get('quantity', ''))))
                items_table.setItem(i, 2, QTableWidgetItem(str(item.get('total_price', ''))))
                items_table.setItem(i, 3, QTableWidgetItem(str(item.get('unit_price', ''))))

            items_table.horizontalHeader().setStretchLastSection(True)
            items_table.resizeColumnsToContents()
            items_layout.addWidget(items_table)
        else:
            no_items_label = QLabel("No line items extracted")
            no_items_label.setStyleSheet("color: #e74c3c; font-style: italic;")
            items_layout.addWidget(no_items_label)

        test_layout.addWidget(items_group)

        # Errors
        if test_results['errors']:
            err_group = QGroupBox("Errors")
            err_layout = QVBoxLayout(err_group)
            for err in test_results['errors']:
                err_label = QLabel(f"• {err}")
                err_label.setStyleSheet("color: #e74c3c;")
                err_layout.addWidget(err_label)
            test_layout.addWidget(err_group)

        test_layout.addStretch()
        tabs.addTab(test_widget, "Test Results")

        # Tab 2: Code Preview
        code_widget = QWidget()
        code_layout = QVBoxLayout(code_widget)
        code_view = QPlainTextEdit()
        code_view.setFont(QFont("Consolas", 10))
        code_view.setPlainText(code)
        code_layout.addWidget(code_view)
        tabs.addTab(code_widget, "Generated Code")

        layout.addWidget(tabs)

        # Buttons
        btn_layout = QHBoxLayout()

        # Warning if validation failed
        if validation_errors or not test_results['can_process']:
            warn_label = QLabel("⚠️ Template has issues - review before saving")
            warn_label.setStyleSheet("color: #f39c12; font-weight: bold;")
            btn_layout.addWidget(warn_label)

        btn_layout.addStretch()

        btn_save = QPushButton("Save Template")
        btn_save.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 10px 20px;")
        btn_save.clicked.connect(lambda: self._save_template(template_name, class_name, code_view.toPlainText(), preview_dialog))
        btn_layout.addWidget(btn_save)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(preview_dialog.reject)
        btn_layout.addWidget(btn_cancel)

        layout.addLayout(btn_layout)

        preview_dialog.exec_()

    def _save_template(self, template_name: str, class_name: str, code: str, preview_dialog: QDialog):
        """Save the template file."""
        templates_dir = Path(__file__).parent / "templates"
        templates_dir.mkdir(exist_ok=True)

        file_path = templates_dir / f"{template_name}.py"

        if file_path.exists():
            result = QMessageBox.question(
                self, "File Exists",
                f"{template_name}.py already exists. Overwrite?",
                QMessageBox.Yes | QMessageBox.No
            )
            if result != QMessageBox.Yes:
                return

        try:
            file_path.write_text(code)

            # Auto-register
            registered = self._auto_register_template(templates_dir, template_name, class_name)

            if registered:
                QMessageBox.information(
                    self, "Template Created",
                    f"Template '{template_name}' has been created and registered!\n\n"
                    f"Click 'Refresh' in the Templates tab to use it."
                )
            else:
                QMessageBox.information(
                    self, "Template Saved",
                    f"Template saved to:\n{file_path}\n\n"
                    f"Auto-registration failed. Please manually add to templates/__init__.py"
                )

            self.template_created.emit(template_name, str(file_path))
            preview_dialog.accept()
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save template: {e}")

    def _auto_register_template(self, templates_dir: Path, template_name: str, class_name: str) -> bool:
        """Auto-register the template in templates/__init__.py."""
        init_file = templates_dir / "__init__.py"
        if not init_file.exists():
            return False

        try:
            content = init_file.read_text()

            # Check if already registered
            if f"'{template_name}'" in content or f'"{template_name}"' in content:
                return True

            # Add import statement
            import_line = f"from .{template_name} import {class_name}\n"

            # Find last import line
            lines = content.split('\n')
            last_import_idx = 0
            for i, line in enumerate(lines):
                if line.startswith('from .') and 'import' in line:
                    last_import_idx = i

            lines.insert(last_import_idx + 1, import_line.rstrip())

            # Add to TEMPLATE_REGISTRY
            registry_entry = f"    '{template_name}': {class_name},"

            new_lines = []
            in_registry = False
            added_entry = False

            for line in lines:
                if 'TEMPLATE_REGISTRY' in line and '{' in line:
                    in_registry = True
                if in_registry and '}' in line and not added_entry:
                    new_lines.append(registry_entry)
                    added_entry = True
                    in_registry = False
                new_lines.append(line)

            init_file.write_text('\n'.join(new_lines))
            return True

        except Exception as e:
            print(f"Auto-register failed: {e}")
            return False
