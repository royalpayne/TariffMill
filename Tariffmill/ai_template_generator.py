"""
AI Template Generator for TariffMill

Allows users to generate invoice templates using AI models (OpenAI, Anthropic, or local models).
The AI analyzes sample invoice text and generates a complete template class.
"""

import os
import re
import json
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QFormLayout, QLineEdit, QTextEdit, QPlainTextEdit,
    QComboBox, QSpinBox, QCheckBox, QFileDialog, QMessageBox,
    QTabWidget, QWidget, QProgressBar, QApplication
)
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont

# Try to import AI libraries
try:
    import openai
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


class AIGeneratorThread(QThread):
    """Background thread for AI template generation."""
    finished = pyqtSignal(str)  # Generated code
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, provider: str, model: str, api_key: str,
                 invoice_text: str, template_name: str, supplier_name: str,
                 country: str, client: str):
        super().__init__()
        self.provider = provider
        self.model = model
        self.api_key = api_key
        self.invoice_text = invoice_text
        self.template_name = template_name
        self.supplier_name = supplier_name
        self.country = country
        self.client = client

    def run(self):
        try:
            self.progress.emit("Preparing prompt...")
            prompt = self._build_prompt()

            self.progress.emit(f"Calling {self.provider} API...")

            if self.provider == "OpenAI":
                result = self._call_openai(prompt)
            elif self.provider == "Anthropic":
                result = self._call_anthropic(prompt)
            elif self.provider == "Ollama (Local)":
                result = self._call_ollama(prompt)
            else:
                raise ValueError(f"Unknown provider: {self.provider}")

            self.progress.emit("Processing response...")
            code = self._extract_code(result)
            self.finished.emit(code)

        except Exception as e:
            self.error.emit(str(e))

    def _build_prompt(self) -> str:
        """Build the prompt for AI template generation."""
        # Truncate invoice text if too long
        invoice_sample = self.invoice_text[:8000] if len(self.invoice_text) > 8000 else self.invoice_text

        prompt = f'''You are an expert Python developer creating invoice parsing templates.
Analyze this sample invoice text and generate a complete Python template class.

SAMPLE INVOICE TEXT:
```
{invoice_sample}
```

REQUIREMENTS:
1. Template Name: {self.template_name}
2. Supplier Name: {self.supplier_name}
3. Country of Origin: {self.country}
4. Client: {self.client}

Generate a Python class that:
1. Inherits from BaseTemplate
2. Has a can_process() method that identifies this supplier's invoices
3. Has extract_invoice_number() to find the invoice number
4. Has extract_project_number() to find PO/project numbers
5. Has extract_line_items() to extract part numbers, quantities, and prices
6. Uses regex patterns appropriate for this invoice format

The class should follow this structure:

```python
"""
{self.supplier_name} Template

Auto-generated template for invoices from {self.supplier_name}.
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""

import re
from typing import List, Dict
from .base_template import BaseTemplate


class {self._to_class_name(self.template_name)}(BaseTemplate):
    """Template for {self.supplier_name} invoices."""

    name = "{self.supplier_name}"
    description = "Invoices from {self.supplier_name}"
    client = "{self.client}"
    version = "1.0.0"
    enabled = True

    extra_columns = ['po_number', 'unit_price', 'description', 'country_origin']

    # Keywords to identify this supplier
    SUPPLIER_KEYWORDS = [
        # Add lowercase keywords here
    ]

    def can_process(self, text: str) -> bool:
        """Check if this is a {self.supplier_name} invoice."""
        text_lower = text.lower()
        for keyword in self.SUPPLIER_KEYWORDS:
            if keyword in text_lower:
                return True
        return False

    def get_confidence_score(self, text: str) -> float:
        """Return confidence score for template matching."""
        if not self.can_process(text):
            return 0.0
        return 0.8

    def extract_invoice_number(self, text: str) -> str:
        """Extract invoice number using regex patterns."""
        patterns = [
            # Add patterns based on the invoice format
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return "UNKNOWN"

    def extract_project_number(self, text: str) -> str:
        """Extract PO/project number."""
        patterns = [
            # Add patterns for PO numbers
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)
        return "UNKNOWN"

    def extract_manufacturer_name(self, text: str) -> str:
        """Return the manufacturer name."""
        return "{self.supplier_name.upper()}"

    def extract_line_items(self, text: str) -> List[Dict]:
        """Extract line items from invoice."""
        items = []
        # Add extraction logic based on the invoice format
        # Look for patterns like: part_number quantity price

        return items

    def post_process_items(self, items: List[Dict]) -> List[Dict]:
        """Post-process - deduplicate and validate."""
        if not items:
            return items

        seen = set()
        unique_items = []

        for item in items:
            key = f"{{item['part_number']}}_{{item['quantity']}}_{{item['total_price']}}"
            if key not in seen:
                seen.add(key)
                # Add country of origin
                item['country_origin'] = '{self.country}'
                unique_items.append(item)

        return unique_items

    def is_packing_list(self, text: str) -> bool:
        """Check if document is only a packing list."""
        text_lower = text.lower()
        if 'packing list' in text_lower and 'invoice' not in text_lower:
            return True
        return False
```

IMPORTANT:
1. Analyze the invoice text carefully to identify the actual patterns used
2. Create specific regex patterns for invoice numbers, PO numbers, and line items
3. The SUPPLIER_KEYWORDS should contain unique identifiers from the invoice
4. The extract_line_items() method should parse the actual table format from the invoice
5. Return ONLY the Python code, no explanations

Generate the complete, working Python template code:
'''
        return prompt

    def _to_class_name(self, template_name: str) -> str:
        """Convert template name to class name."""
        return ''.join(word.title() for word in template_name.split('_')) + 'Template'

    def _call_openai(self, prompt: str) -> str:
        """Call OpenAI API."""
        if not HAS_OPENAI:
            raise ImportError("openai package not installed. Run: pip install openai")

        client = openai.OpenAI(api_key=self.api_key)
        response = client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": "You are an expert Python developer specializing in invoice parsing and OCR templates."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=4000,
            temperature=0.3
        )
        return response.choices[0].message.content

    def _call_anthropic(self, prompt: str) -> str:
        """Call Anthropic API."""
        if not HAS_ANTHROPIC:
            raise ImportError("anthropic package not installed. Run: pip install anthropic")

        client = anthropic.Anthropic(api_key=self.api_key)
        response = client.messages.create(
            model=self.model,
            max_tokens=4000,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.content[0].text

    def _call_ollama(self, prompt: str) -> str:
        """Call local Ollama API."""
        import urllib.request
        import json

        url = "http://localhost:11434/api/generate"
        data = {
            "model": self.model,
            "prompt": prompt,
            "stream": False
        }

        req = urllib.request.Request(
            url,
            data=json.dumps(data).encode('utf-8'),
            headers={'Content-Type': 'application/json'}
        )

        try:
            with urllib.request.urlopen(req, timeout=120) as response:
                result = json.loads(response.read().decode('utf-8'))
                return result.get('response', '')
        except Exception as e:
            raise ConnectionError(f"Failed to connect to Ollama: {e}\n\nMake sure Ollama is running (ollama serve)")

    def _extract_code(self, response: str) -> str:
        """Extract Python code from AI response."""
        # Try to find code block
        code_match = re.search(r'```python\s*(.*?)\s*```', response, re.DOTALL)
        if code_match:
            return code_match.group(1).strip()

        # Try without language specifier
        code_match = re.search(r'```\s*(.*?)\s*```', response, re.DOTALL)
        if code_match:
            return code_match.group(1).strip()

        # Return the whole response if no code blocks found
        # but try to clean it up
        lines = response.split('\n')
        code_lines = []
        in_code = False

        for line in lines:
            if line.strip().startswith('"""') or line.strip().startswith('import ') or line.strip().startswith('from '):
                in_code = True
            if in_code:
                code_lines.append(line)

        if code_lines:
            return '\n'.join(code_lines)

        return response


class AITemplateGeneratorDialog(QDialog):
    """
    Dialog for generating invoice templates using AI models.

    Supports:
    - OpenAI (GPT-4, GPT-3.5)
    - Anthropic (Claude)
    - Ollama (local models)
    """

    template_created = pyqtSignal(str, str)  # template_name, file_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.generator_thread = None
        self.invoice_text = ""

        self.setWindowTitle("AI Template Generator")
        self.setMinimumSize(900, 700)
        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """Build the dialog UI."""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header
        header = QLabel("AI Template Generator")
        header.setFont(QFont("Arial", 16, QFont.Bold))
        header.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(header)

        desc = QLabel(
            "Use AI to automatically generate invoice templates from sample invoices. "
            "Load a PDF or paste invoice text, then let AI create the extraction patterns."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #7f8c8d; margin-bottom: 15px;")
        layout.addWidget(desc)

        # AI Provider Settings
        provider_group = QGroupBox("AI Provider")
        provider_layout = QFormLayout()

        self.provider_combo = QComboBox()
        providers = []
        if HAS_OPENAI:
            providers.append("OpenAI")
        if HAS_ANTHROPIC:
            providers.append("Anthropic")
        providers.append("Ollama (Local)")  # Always available

        if not providers:
            providers = ["Ollama (Local)"]

        self.provider_combo.addItems(providers)
        self.provider_combo.currentTextChanged.connect(self.on_provider_changed)
        provider_layout.addRow("Provider:", self.provider_combo)

        self.model_combo = QComboBox()
        self.model_combo.setEditable(True)
        provider_layout.addRow("Model:", self.model_combo)

        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.api_key_edit.setPlaceholderText("Enter API key (not needed for Ollama)")
        provider_layout.addRow("API Key:", self.api_key_edit)

        provider_group.setLayout(provider_layout)
        layout.addWidget(provider_group)

        # Update models for initial provider
        self.on_provider_changed(self.provider_combo.currentText())

        # Invoice Input
        input_group = QGroupBox("Sample Invoice")
        input_layout = QVBoxLayout()

        # File selection row
        file_row = QHBoxLayout()
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("Select a PDF invoice or paste text below...")
        self.pdf_path_edit.setReadOnly(True)
        file_row.addWidget(self.pdf_path_edit, stretch=1)

        browse_btn = QPushButton("Browse PDF...")
        browse_btn.clicked.connect(self.browse_pdf)
        file_row.addWidget(browse_btn)

        input_layout.addLayout(file_row)

        self.invoice_text_edit = QPlainTextEdit()
        self.invoice_text_edit.setPlaceholderText(
            "Paste invoice text here, or load from PDF above.\n\n"
            "The AI will analyze this text to create extraction patterns."
        )
        self.invoice_text_edit.setFont(QFont("Courier New", 9))
        input_layout.addWidget(self.invoice_text_edit)

        input_group.setLayout(input_layout)
        layout.addWidget(input_group, stretch=1)

        # Template Settings
        settings_group = QGroupBox("Template Settings")
        settings_layout = QFormLayout()

        self.template_name_edit = QLineEdit()
        self.template_name_edit.setPlaceholderText("e.g., acme_corp (lowercase with underscores)")
        settings_layout.addRow("Template Name:", self.template_name_edit)

        self.supplier_name_edit = QLineEdit()
        self.supplier_name_edit.setPlaceholderText("e.g., Acme Corporation")
        settings_layout.addRow("Supplier Name:", self.supplier_name_edit)

        self.client_edit = QLineEdit()
        self.client_edit.setPlaceholderText("e.g., Sigma Corporation")
        settings_layout.addRow("Client:", self.client_edit)

        self.country_edit = QLineEdit()
        self.country_edit.setPlaceholderText("e.g., CHINA, INDIA, USA")
        settings_layout.addRow("Country of Origin:", self.country_edit)

        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)

        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # Generated Code Preview
        preview_group = QGroupBox("Generated Template (Preview)")
        preview_layout = QVBoxLayout()

        self.code_preview = QPlainTextEdit()
        self.code_preview.setReadOnly(True)
        self.code_preview.setFont(QFont("Courier New", 9))
        self.code_preview.setStyleSheet("background-color: #f8f9fa;")
        self.code_preview.setPlaceholderText("Generated template code will appear here...")
        preview_layout.addWidget(self.code_preview)

        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group, stretch=1)

        # Action buttons
        btn_layout = QHBoxLayout()

        self.generate_btn = QPushButton("Generate Template")
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-weight: bold;
                padding: 10px 25px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_template)
        btn_layout.addWidget(self.generate_btn)

        self.save_btn = QPushButton("Save Template")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                padding: 10px 25px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
        """)
        self.save_btn.clicked.connect(self.save_template)
        self.save_btn.setEnabled(False)
        btn_layout.addWidget(self.save_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def on_provider_changed(self, provider: str):
        """Update model list when provider changes."""
        self.model_combo.clear()

        if provider == "OpenAI":
            self.model_combo.addItems(["gpt-4o", "gpt-4-turbo", "gpt-4", "gpt-3.5-turbo"])
            self.api_key_edit.setEnabled(True)
            self.api_key_edit.setPlaceholderText("Enter OpenAI API key")
            # Try to load from environment
            if os.environ.get('OPENAI_API_KEY'):
                self.api_key_edit.setText(os.environ['OPENAI_API_KEY'])
        elif provider == "Anthropic":
            self.model_combo.addItems(["claude-sonnet-4-20250514", "claude-3-5-sonnet-20241022", "claude-3-opus-20240229", "claude-3-haiku-20240307"])
            self.api_key_edit.setEnabled(True)
            self.api_key_edit.setPlaceholderText("Enter Anthropic API key")
            if os.environ.get('ANTHROPIC_API_KEY'):
                self.api_key_edit.setText(os.environ['ANTHROPIC_API_KEY'])
        elif provider == "Ollama (Local)":
            self.model_combo.addItems(["llama3.1", "llama3", "codellama", "mistral", "mixtral", "deepseek-coder"])
            self.api_key_edit.setEnabled(False)
            self.api_key_edit.setPlaceholderText("Not needed for local Ollama")
            self.api_key_edit.clear()

    def browse_pdf(self):
        """Browse for PDF file and extract text."""
        if not HAS_PDFPLUMBER:
            QMessageBox.warning(
                self, "Missing Dependency",
                "pdfplumber is required to load PDFs.\n\n"
                "Install with: pip install pdfplumber"
            )
            return

        path, _ = QFileDialog.getOpenFileName(
            self, "Select Invoice PDF",
            str(Path.home()),
            "PDF Files (*.pdf);;All Files (*.*)"
        )

        if not path:
            return

        try:
            self.pdf_path_edit.setText(path)

            # Extract text from PDF
            text_parts = []
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages[:5]:  # First 5 pages
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)

            full_text = '\n\n'.join(text_parts)
            self.invoice_text_edit.setPlainText(full_text)

            # Try to auto-detect supplier name
            self._auto_detect_supplier(full_text)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract PDF text:\n{e}")

    def _auto_detect_supplier(self, text: str):
        """Try to auto-detect supplier name from invoice text."""
        # Look for common patterns
        lines = text.split('\n')

        # First few non-empty lines often contain company name
        for line in lines[:10]:
            line = line.strip()
            if len(line) > 5 and len(line) < 60:
                # Skip lines that look like addresses or dates
                if re.match(r'^[\d\s\-/]+$', line):  # Just numbers
                    continue
                if re.search(r'\d{4}', line):  # Contains year
                    continue
                if '@' in line or 'www.' in line.lower():  # Email/web
                    continue

                # This might be a company name
                if not self.supplier_name_edit.text():
                    self.supplier_name_edit.setText(line)
                    # Generate template name
                    template_name = re.sub(r'[^a-z0-9]+', '_', line.lower()).strip('_')
                    self.template_name_edit.setText(template_name[:30])
                break

    def generate_template(self):
        """Start AI template generation."""
        # Validate inputs
        invoice_text = self.invoice_text_edit.toPlainText().strip()
        if not invoice_text:
            QMessageBox.warning(self, "Missing Input", "Please provide invoice text or load a PDF.")
            return

        template_name = self.template_name_edit.text().strip()
        if not template_name:
            QMessageBox.warning(self, "Missing Input", "Please enter a template name.")
            return

        if not re.match(r'^[a-z][a-z0-9_]*$', template_name):
            QMessageBox.warning(
                self, "Invalid Name",
                "Template name must be lowercase, start with a letter, "
                "and contain only letters, numbers, and underscores."
            )
            return

        supplier_name = self.supplier_name_edit.text().strip()
        if not supplier_name:
            QMessageBox.warning(self, "Missing Input", "Please enter a supplier name.")
            return

        provider = self.provider_combo.currentText()
        api_key = self.api_key_edit.text().strip()

        if provider in ["OpenAI", "Anthropic"] and not api_key:
            QMessageBox.warning(self, "Missing API Key", f"Please enter your {provider} API key.")
            return

        # Start generation
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFormat("Generating...")
        self.generate_btn.setEnabled(False)
        self.save_btn.setEnabled(False)

        self.generator_thread = AIGeneratorThread(
            provider=provider,
            model=self.model_combo.currentText(),
            api_key=api_key,
            invoice_text=invoice_text,
            template_name=template_name,
            supplier_name=supplier_name,
            country=self.country_edit.text().strip() or "UNKNOWN",
            client=self.client_edit.text().strip() or "Universal"
        )
        self.generator_thread.finished.connect(self.on_generation_complete)
        self.generator_thread.error.connect(self.on_generation_error)
        self.generator_thread.progress.connect(self.on_progress)
        self.generator_thread.start()

    def on_progress(self, message: str):
        """Update progress message."""
        self.progress_bar.setFormat(message)

    def on_generation_complete(self, code: str):
        """Handle successful generation."""
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        self.save_btn.setEnabled(True)

        self.code_preview.setPlainText(code)

        QMessageBox.information(
            self, "Generation Complete",
            "Template generated successfully!\n\n"
            "Review the code in the preview, then click 'Save Template' to save it."
        )

    def on_generation_error(self, error: str):
        """Handle generation error."""
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)

        QMessageBox.critical(
            self, "Generation Error",
            f"Failed to generate template:\n\n{error}"
        )

    def save_template(self):
        """Save the generated template."""
        code = self.code_preview.toPlainText().strip()
        if not code:
            QMessageBox.warning(self, "No Code", "No template code to save.")
            return

        template_name = self.template_name_edit.text().strip()

        # Determine templates directory
        templates_dir = Path(__file__).parent / 'templates'
        if not templates_dir.exists():
            templates_dir = Path(__file__).parent.parent / 'Tariffmill' / 'templates'

        file_path = templates_dir / f"{template_name}.py"

        # Check if file already exists
        if file_path.exists():
            reply = QMessageBox.question(
                self, "File Exists",
                f"Template '{template_name}' already exists.\n\nOverwrite?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        # Save the file
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(code)

            QMessageBox.information(
                self, "Template Saved",
                f"Template saved successfully!\n\n"
                f"File: {file_path}\n\n"
                f"The template will be available after refreshing templates."
            )

            # Emit signal
            self.template_created.emit(template_name, str(file_path))

            # Save settings
            self.save_settings()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save template:\n{e}")

    def load_settings(self):
        """Load saved settings (API keys, preferred provider)."""
        # Try to load from environment variables first
        if os.environ.get('OPENAI_API_KEY') and self.provider_combo.currentText() == "OpenAI":
            self.api_key_edit.setText(os.environ['OPENAI_API_KEY'])
        elif os.environ.get('ANTHROPIC_API_KEY') and self.provider_combo.currentText() == "Anthropic":
            self.api_key_edit.setText(os.environ['ANTHROPIC_API_KEY'])

    def save_settings(self):
        """Save settings for next time."""
        # Could save to a config file, but for now just keep in memory
        pass


def main():
    """Test the dialog standalone."""
    import sys
    app = QApplication(sys.argv)
    dialog = AITemplateGeneratorDialog()
    dialog.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()