"""
Auto Template Generator Dialog for TariffMill
Provides a GUI interface for analyzing PDFs and generating invoice templates.
"""

import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QTextEdit, QFileDialog, QGroupBox, QFormLayout,
    QMessageBox, QProgressBar, QSplitter, QWidget, QTabWidget
)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont

# Base directory for templates
BASE_DIR = Path(__file__).parent


class AnalysisWorker(QThread):
    """Worker thread for PDF analysis to keep UI responsive."""
    finished = pyqtSignal(object)  # Emits the analysis result
    error = pyqtSignal(str)  # Emits error message
    progress = pyqtSignal(str)  # Emits progress messages

    def __init__(self, pdf_path: str, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path

    def run(self):
        try:
            self.progress.emit("Loading PDF...")
            from template_generator import TemplateGenerator
            generator = TemplateGenerator()

            self.progress.emit("Analyzing invoice patterns...")
            analysis = generator.analyze_pdf(self.pdf_path)

            self.progress.emit("Analysis complete")
            self.finished.emit((generator, analysis))

        except ImportError as e:
            self.error.emit(f"Missing dependency: {e}\n\nInstall pdfplumber: pip install pdfplumber")
        except Exception as e:
            self.error.emit(str(e))


class AutoTemplateGeneratorDialog(QDialog):
    """
    Auto Template Generator Dialog.
    Analyzes PDF invoices and generates template code automatically.
    """

    template_created = pyqtSignal(str, str)  # template_name, file_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Auto Template Generator")
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint)
        self.setMinimumSize(900, 700)

        self.generator = None
        self.analysis = None
        self.current_pdf = None

        self._setup_ui()

    def _setup_ui(self):
        """Set up the dialog UI."""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header
        header = QLabel("Auto Template Generator")
        header.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(header)

        desc = QLabel(
            "Analyze a sample PDF invoice to automatically detect patterns and generate template code.\n"
            "Select a PDF, analyze it, review the detected patterns, then generate the template."
        )
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # Step 1: PDF Selection
        pdf_group = QGroupBox("Step 1: Select Sample PDF")
        pdf_layout = QHBoxLayout()

        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("Select a sample invoice PDF to analyze...")
        self.pdf_path_edit.setReadOnly(True)
        pdf_layout.addWidget(self.pdf_path_edit, 1)

        self.btn_browse = QPushButton("Browse...")
        self.btn_browse.clicked.connect(self._browse_pdf)
        pdf_layout.addWidget(self.btn_browse)

        self.btn_analyze = QPushButton("Analyze PDF")
        self.btn_analyze.setEnabled(False)
        self.btn_analyze.clicked.connect(self._analyze_pdf)
        pdf_layout.addWidget(self.btn_analyze)

        pdf_group.setLayout(pdf_layout)
        layout.addWidget(pdf_group)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # Main content area with tabs
        self.tabs = QTabWidget()

        # Tab 1: Analysis Results
        analysis_widget = QWidget()
        analysis_layout = QVBoxLayout(analysis_widget)

        self.analysis_text = QTextEdit()
        self.analysis_text.setReadOnly(True)
        self.analysis_text.setFont(QFont("Consolas", 10))
        self.analysis_text.setPlaceholderText("Analysis results will appear here after analyzing a PDF...")
        analysis_layout.addWidget(self.analysis_text)

        self.tabs.addTab(analysis_widget, "Analysis Results")

        # Tab 2: Generated Code
        code_widget = QWidget()
        code_layout = QVBoxLayout(code_widget)

        self.code_text = QTextEdit()
        self.code_text.setFont(QFont("Consolas", 10))
        self.code_text.setPlaceholderText("Generated template code will appear here...")
        code_layout.addWidget(self.code_text)

        self.tabs.addTab(code_widget, "Generated Code")

        # Tab 3: Sample Text
        sample_widget = QWidget()
        sample_layout = QVBoxLayout(sample_widget)

        self.sample_text = QTextEdit()
        self.sample_text.setReadOnly(True)
        self.sample_text.setFont(QFont("Consolas", 9))
        self.sample_text.setPlaceholderText("Extracted text from the PDF will appear here...")
        sample_layout.addWidget(self.sample_text)

        self.tabs.addTab(sample_widget, "PDF Text")

        layout.addWidget(self.tabs, 1)

        # Step 2: Template Configuration
        config_group = QGroupBox("Step 2: Configure Template")
        config_layout = QFormLayout()

        self.template_name_edit = QLineEdit()
        self.template_name_edit.setPlaceholderText("e.g., acme_corp_invoices")
        self.template_name_edit.textChanged.connect(self._update_generate_button)
        config_layout.addRow("Template Name:", self.template_name_edit)

        self.class_name_edit = QLineEdit()
        self.class_name_edit.setPlaceholderText("Auto-generated from template name")
        config_layout.addRow("Class Name (optional):", self.class_name_edit)

        config_group.setLayout(config_layout)
        layout.addWidget(config_group)

        # Action buttons
        buttons_layout = QHBoxLayout()

        self.btn_generate = QPushButton("Generate Template Code")
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self._generate_template)
        buttons_layout.addWidget(self.btn_generate)

        self.btn_save = QPushButton("Save Template")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self._save_template)
        buttons_layout.addWidget(self.btn_save)

        buttons_layout.addStretch()

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.close)
        buttons_layout.addWidget(btn_close)

        layout.addLayout(buttons_layout)

        # Apply styles
        self._apply_styles()

    def _apply_styles(self):
        """Apply consistent styling to the dialog."""
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                min-width: 100px;
                padding: 6px 12px;
            }
            QPushButton:disabled {
                background-color: #e0e0e0;
                color: #888;
            }
            QTextEdit {
                border: 1px solid #ccc;
                border-radius: 3px;
            }
        """)

    def _browse_pdf(self):
        """Open file dialog to select a PDF."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Sample PDF Invoice",
            str(Path.home()),
            "PDF Files (*.pdf);;All Files (*)"
        )

        if file_path:
            self.pdf_path_edit.setText(file_path)
            self.current_pdf = file_path
            self.btn_analyze.setEnabled(True)

            # Auto-suggest template name from filename
            filename = Path(file_path).stem
            # Clean up filename for use as template name
            clean_name = filename.lower().replace(' ', '_').replace('-', '_')
            clean_name = ''.join(c for c in clean_name if c.isalnum() or c == '_')
            self.template_name_edit.setText(clean_name)

    def _analyze_pdf(self):
        """Analyze the selected PDF."""
        if not self.current_pdf:
            return

        # Show progress
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.setFormat("Analyzing...")
        self.btn_analyze.setEnabled(False)
        self.btn_browse.setEnabled(False)

        # Clear previous results
        self.analysis_text.clear()
        self.code_text.clear()
        self.sample_text.clear()

        # Run analysis in background thread
        self.worker = AnalysisWorker(self.current_pdf)
        self.worker.finished.connect(self._on_analysis_complete)
        self.worker.error.connect(self._on_analysis_error)
        self.worker.progress.connect(self._on_analysis_progress)
        self.worker.start()

    def _on_analysis_progress(self, message: str):
        """Handle progress updates from worker."""
        self.progress_bar.setFormat(message)

    def _on_analysis_complete(self, result):
        """Handle completed analysis."""
        self.generator, self.analysis = result

        # Hide progress
        self.progress_bar.setVisible(False)
        self.btn_analyze.setEnabled(True)
        self.btn_browse.setEnabled(True)

        # Display analysis results
        self._display_analysis()

        # Show sample text
        if self.analysis and self.analysis.raw_text:
            self.sample_text.setPlainText(self.analysis.raw_text[:10000])

        # Enable generate button
        self._update_generate_button()

        # Switch to analysis tab
        self.tabs.setCurrentIndex(0)

    def _on_analysis_error(self, error_msg: str):
        """Handle analysis error."""
        self.progress_bar.setVisible(False)
        self.btn_analyze.setEnabled(True)
        self.btn_browse.setEnabled(True)

        QMessageBox.critical(self, "Analysis Error", f"Failed to analyze PDF:\n\n{error_msg}")

    def _display_analysis(self):
        """Display the analysis results in a formatted way."""
        if not self.analysis:
            return

        a = self.analysis
        lines = []

        lines.append("=" * 60)
        lines.append("INVOICE ANALYSIS SUMMARY")
        lines.append("=" * 60)
        lines.append("")

        lines.append(f"Supplier/Company: {a.supplier_name or 'Unknown'}")
        if a.supplier_indicators:
            lines.append(f"Indicators: {', '.join(a.supplier_indicators)}")
        lines.append("")

        if a.invoice_number_pattern:
            lines.append("INVOICE NUMBER PATTERN:")
            lines.append(f"  Pattern: {a.invoice_number_pattern.pattern}")
            lines.append(f"  Samples: {a.invoice_number_pattern.sample_matches}")
            lines.append(f"  Confidence: {a.invoice_number_pattern.confidence:.0%}")
            lines.append("")

        if a.project_number_pattern:
            lines.append("PROJECT/PO NUMBER PATTERN:")
            lines.append(f"  Pattern: {a.project_number_pattern.pattern}")
            lines.append(f"  Samples: {a.project_number_pattern.sample_matches}")
            lines.append(f"  Confidence: {a.project_number_pattern.confidence:.0%}")
            lines.append("")

        if a.line_item_pattern:
            lines.append("LINE ITEM PATTERN:")
            lines.append(f"  Pattern: {a.line_item_pattern.pattern}")
            lines.append(f"  Fields: {', '.join(a.line_item_pattern.field_names)}")
            lines.append(f"  Confidence: {a.line_item_pattern.confidence:.0%}")
            if a.line_item_pattern.sample_matches:
                lines.append("  Sample matches:")
                for match in a.line_item_pattern.sample_matches[:3]:
                    lines.append(f"    {match}")
            lines.append("")

        if a.extra_fields:
            lines.append("ADDITIONAL FIELDS DETECTED:")
            for name, pattern in a.extra_fields.items():
                lines.append(f"  {name}: {pattern.sample_matches[:3]}")
            lines.append("")

        lines.append("=" * 60)
        lines.append("SAMPLE LINES FROM INVOICE:")
        lines.append("=" * 60)
        for line in a.sample_lines[:15]:
            lines.append(f"  {line[:100]}")

        self.analysis_text.setPlainText('\n'.join(lines))

    def _update_generate_button(self):
        """Update the generate button state."""
        has_analysis = self.analysis is not None
        has_name = bool(self.template_name_edit.text().strip())
        self.btn_generate.setEnabled(has_analysis and has_name)

    def _generate_template(self):
        """Generate template code from the analysis."""
        if not self.generator or not self.analysis:
            return

        template_name = self.template_name_edit.text().strip()
        class_name = self.class_name_edit.text().strip() or None

        try:
            code = self.generator.generate_template(
                template_name=template_name,
                class_name=class_name
            )

            self.code_text.setPlainText(code)
            self.btn_save.setEnabled(True)

            # Switch to code tab
            self.tabs.setCurrentIndex(1)

        except Exception as e:
            QMessageBox.critical(self, "Generation Error", f"Failed to generate template:\n\n{e}")

    def _save_template(self):
        """Save the generated template to the templates directory."""
        template_name = self.template_name_edit.text().strip()
        if not template_name:
            QMessageBox.warning(self, "Missing Name", "Please enter a template name.")
            return

        code = self.code_text.toPlainText()
        if not code:
            QMessageBox.warning(self, "No Code", "Please generate template code first.")
            return

        # Validate name
        clean_name = template_name.lower().replace(' ', '_').replace('-', '_')
        if not clean_name.replace('_', '').isalnum():
            QMessageBox.warning(
                self, "Invalid Name",
                "Template name must contain only letters, numbers, and underscores."
            )
            return

        # Check for existing template
        templates_dir = BASE_DIR / "templates"
        file_path = templates_dir / f"{clean_name}.py"

        if file_path.exists():
            reply = QMessageBox.question(
                self, "Template Exists",
                f"A template named '{clean_name}' already exists.\n\nDo you want to overwrite it?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        # Save the template
        try:
            with open(file_path, 'w') as f:
                f.write(code)

            QMessageBox.information(
                self, "Template Saved",
                f"Template saved successfully!\n\n"
                f"File: {file_path}\n\n"
                "The template will be available after refreshing the templates list."
            )

            # Emit signal to notify parent
            self.template_created.emit(clean_name, str(file_path))

        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save template:\n\n{e}")
