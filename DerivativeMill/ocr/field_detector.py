"""
Field Detection and Supplier Templates

Provides supplier-specific templates for extracting Part Number and Value fields
from OCR-extracted text using pattern matching.
"""

import re
import json
from pathlib import Path


class SupplierTemplate:
    """
    Template for extracting Part Number and Value from invoice text.

    Each supplier may have a different invoice format. This class defines
    regex patterns and field positions for a specific supplier.
    """

    def __init__(self, supplier_name, patterns=None, field_positions=None):
        """
        Initialize a supplier template.

        Args:
            supplier_name (str): Name of supplier (e.g., "ACME Corp")
            patterns (dict): Regex patterns for field detection
            field_positions (dict): Optional positional info for fields
        """
        self.supplier_name = supplier_name
        self.patterns = patterns or self._default_patterns()
        self.field_positions = field_positions or {}

    def _default_patterns(self):
        """Default patterns that work for most invoices."""
        return {
            'part_number_header': r'(part\s*(?:number|num|no|#|code|id)|sku|product\s*(?:number|id|code)|item\s*(?:number|code))',
            'part_number_value': r'([A-Z0-9\-_\.]{3,25})',
            'value_header': r'(price|unit\s*price|value|amount|cost|rate|total|invoice|qty|quantity)',
            'value_pattern': r'\$?\s*(\d{1,10}(?:[,\.]?\d{1,3})*(?:\.\d{2})?)',
            'quantity_pattern': r'qty:?\s*(\d+)',
            'description': r'(description|desc|item\s*description)',
        }

    def extract(self, text):
        """
        Extract Part Number and Value entries from OCR text.

        Args:
            text (str): OCR-extracted text from invoice

        Returns:
            list: List of dicts with {'part_number': '...', 'value': '...', 'quantity': ...}

        Raises:
            Exception: If extraction fails
        """
        try:
            lines = text.split('\n')
            extracted_data = []

            # Find header line (contains "Part Number" or "Value")
            header_idx = self._find_header_line(lines)

            if header_idx is None:
                # Fallback: extract without header detection
                return self._extract_without_header(text)

            # Extract data lines after header
            part_col = None
            value_col = None

            for idx in range(header_idx + 1, len(lines)):
                line = lines[idx].strip()

                if not line:
                    continue

                # Try to extract Part Number and Value from this line
                part_num = self._extract_part_number(line)
                value = self._extract_value(line)

                # Only include if BOTH part number AND value are present
                # Also validate that value looks like a price (>= 1.0)
                if part_num and value and value >= 1.0:
                    extracted_data.append({
                        'part_number': part_num,
                        'value': value,
                        'raw_line': line
                    })

            return extracted_data

        except Exception as e:
            raise Exception(f"Field extraction error: {str(e)}")

    def _find_header_line(self, lines):
        """Find the line containing column headers."""
        header_pattern = self.patterns['part_number_header']

        for idx, line in enumerate(lines):
            if re.search(header_pattern, line, re.IGNORECASE):
                return idx

        return None

    def _extract_part_number(self, text):
        """Extract part number from a line of text."""
        # Look for patterns like "ABC-123" or "SKU12345"
        pattern = self.patterns.get('part_number_value', '')

        # Skip extraction if pattern is empty
        if not pattern or not pattern.strip():
            return None

        try:
            match = re.search(pattern, text)
            if match:
                # Handle patterns with or without capturing groups
                part_num = None
                if match.groups():
                    part_num = match.group(1).strip()
                else:
                    part_num = match.group(0).strip()

                # Additional validation: part numbers in invoices usually:
                # 1. Contain BOTH letters AND numbers (or special chars)
                # 2. Are at least 4 characters (filters out short abbreviations)
                # 3. Are NOT just dates (YYYY-MMDD format)
                if part_num and len(part_num) >= 4:
                    has_letter = any(c.isalpha() for c in part_num)
                    has_number = any(c.isdigit() for c in part_num)
                    has_special = any(c in '-_.' for c in part_num)

                    # Must have letters (not just numbers or dates)
                    # And must have special chars or numbers (to make it look like a part code)
                    is_likely_date = re.match(r'^\d{4}-\d{2,4}$', part_num)

                    if has_letter and (has_special or has_number) and not is_likely_date:
                        return part_num

        except (IndexError, AttributeError, re.error):
            pass

        return None

    def _extract_value(self, text):
        """Extract numeric value (price/amount) from text.

        For invoice lines, the actual price/total is usually AFTER the part number
        and NEAR THE END of the line, not at the beginning (which often has dates).

        Returns the value as a float, preserving decimal precision.
        """
        pattern = self.patterns.get('value_pattern', '')

        # Skip extraction if pattern is empty
        if not pattern or not pattern.strip():
            return None

        try:
            matches = re.findall(pattern, text)

            if matches:
                # Find numeric values that look like prices
                # Skip values that look like dates (4 digits like 2025)
                # Keep track of both the float value and original string
                valid_values = []  # List of (float_value, original_string) tuples

                for i, match in enumerate(matches):
                    # Keep original for later reconstruction of decimals
                    original = match
                    cleaned = match.replace(',', '').replace(' ', '')

                    # Skip if this looks like a date/year (4-digit number at start of text)
                    if re.match(r'^\d{4}$', cleaned):
                        continue  # Skip year-like values (2025, etc.)

                    try:
                        val = float(cleaned)
                        # Only consider values that look like prices
                        # Avoid tiny quantities (< 1) and suspicious date-like patterns
                        if val > 1.0 and val < 1000000:  # Reasonable price range
                            valid_values.append((val, original))
                    except ValueError:
                        continue

                if valid_values:
                    # Return the LAST (rightmost) valid value
                    # Convert back to float to ensure numeric consistency
                    # The float will preserve decimal places when converted to string
                    return float(valid_values[-1][1].replace(',', '').replace(' ', ''))

        except (IndexError, AttributeError, re.error):
            return None

        return None

    def _extract_without_header(self, text):
        """Fallback extraction without header detection.

        For invoices without clear headers, intelligently filter to extract
        only actual line item data, not metadata/header information.
        """
        lines = text.split('\n')
        extracted_data = []

        for line in lines:
            line = line.strip()

            if not line or len(line) < 5:
                continue

            # Skip lines that are clearly metadata/headers
            if self._is_metadata_line(line):
                continue

            part_num = self._extract_part_number(line)
            value = self._extract_value(line)

            # Only add lines that have BOTH a part number AND a value
            # AND the value is reasonably formatted (not just a random number)
            if part_num and value:
                # Additional check: value should look like a price (typically >= 1.0)
                # This filters out invoice numbers, postal codes, etc.
                if value >= 1.0:
                    extracted_data.append({
                        'part_number': part_num,
                        'value': value,
                        'raw_line': line
                    })

        return extracted_data

    def _is_metadata_line(self, line):
        """Check if a line is metadata/header, not actual data.

        Very conservative approach - only filter obvious header/footer lines,
        not product descriptions that might contain certain keywords.
        """
        line_lower = line.lower()

        # Only filter lines that are CLEARLY metadata headers/footers
        # Avoid filtering product descriptions that happen to contain common words
        strict_metadata_patterns = [
            # Invoice headers
            r'^\s*commercial invoice',
            r'^\s*invoice no:?',
            r'^\s*shipper\s*:',
            r'^\s*consignee\s*:',
            r'^\s*buyer\s*:',
            r'^\s*destination\s*:',
            r'^\s*country of',
            r'^\s*port of',
            # Company names (too specific, usually at start of lines)
            r'^\s*r\.\s*b\.\s*agarwalla',
            r'^\s*masonry supply',
            # Address components
            r'^\s*\d+\s*,\s*[a-z\s]+\s*road',
            r'^\s*[a-z\s]+\s*,\s*\d+',
            # Footer/declaration lines
            r'^\s*we\s*(?:declare|certify)',
            r'^\s*(?:for|on behalf of)',
            r'^\s*stuffing point',
            r'^\s*iec number',
            r'^\s*total\s*(?:packages|weight)',
            r'^\s*bill of lading',
        ]

        # Check strict patterns with regex
        for pattern in strict_metadata_patterns:
            if re.search(pattern, line_lower):
                return True

        return False

    def to_dict(self):
        """Serialize template to dict for JSON storage."""
        return {
            'supplier_name': self.supplier_name,
            'patterns': self.patterns,
            'field_positions': self.field_positions,
        }

    @classmethod
    def from_dict(cls, data):
        """Deserialize template from dict."""
        return cls(
            supplier_name=data['supplier_name'],
            patterns=data.get('patterns'),
            field_positions=data.get('field_positions'),
        )


class TemplateManager:
    """
    Manages supplier templates for field extraction.

    Stores and retrieves supplier-specific extraction templates.
    """

    def __init__(self, templates_dir=None):
        """
        Initialize template manager.

        Args:
            templates_dir (str): Directory to store/load templates (default: ./ocr/templates/)
        """
        if templates_dir is None:
            templates_dir = Path(__file__).parent / 'templates'

        self.templates_dir = Path(templates_dir)
        self.templates_dir.mkdir(exist_ok=True)
        self.templates = {}
        self._load_all_templates()

    def _load_all_templates(self):
        """Load all templates from disk."""
        for template_file in self.templates_dir.glob('*.json'):
            try:
                with open(template_file) as f:
                    data = json.load(f)
                    template = SupplierTemplate.from_dict(data)
                    self.templates[template.supplier_name] = template
            except Exception as e:
                print(f"Error loading template {template_file}: {e}")

    def get_template(self, supplier_name):
        """
        Get a template by supplier name.

        Returns default template if supplier not found.

        Args:
            supplier_name (str): Name of supplier

        Returns:
            SupplierTemplate: The template (or default)
        """
        if supplier_name in self.templates:
            return self.templates[supplier_name]

        # Return default template
        return SupplierTemplate('default')

    def save_template(self, template):
        """
        Save a template to disk.

        Args:
            template (SupplierTemplate): Template to save
        """
        template_file = self.templates_dir / f"{template.supplier_name}.json"

        with open(template_file, 'w') as f:
            json.dump(template.to_dict(), f, indent=2)

        self.templates[template.supplier_name] = template

    def list_templates(self):
        """
        List all available templates.

        Returns:
            list: List of supplier names
        """
        return list(self.templates.keys())


# Global template manager instance
_template_manager = None


def get_template_manager():
    """Get or create the global template manager instance."""
    global _template_manager

    if _template_manager is None:
        _template_manager = TemplateManager()

    return _template_manager


def extract_fields_from_text(text, supplier_name='default'):
    """
    Extract Part Number and Value fields from OCR text.

    Args:
        text (str): OCR-extracted text from invoice
        supplier_name (str): Name of supplier (for template selection)

    Returns:
        list: List of dicts with extracted fields

    Raises:
        Exception: If extraction fails
    """
    manager = get_template_manager()
    template = manager.get_template(supplier_name)
    return template.extract(text)
