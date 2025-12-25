"""
Sigma_Shaanxi Template

Auto-generated template for invoices from Sigma_Shaanxi.
Uses SmartExtractor for reliable extraction with supplier-specific identification.

Generated: 2025-12-24 00:06:18
"""

import re
from typing import List, Dict
from .base_template import BaseTemplate

import sys
from pathlib import Path

parent_dir = Path(__file__).parent.parent
if str(parent_dir) not in sys.path:
    sys.path.insert(0, str(parent_dir))

try:
    from smart_extractor import SmartExtractor
except ImportError:
    try:
        from Tariffmill.smart_extractor import SmartExtractor
    except ImportError:
        SmartExtractor = None


class SmartShaanxiTemplateTemplate(BaseTemplate):
    """
    Template for Sigma_Shaanxi invoices.
    Uses SmartExtractor for line item extraction.
    """

    name = "Sigma_Shaanxi"
    description = "Invoices from Sigma_Shaanxi"
    client = "Sigma Corporation"
    version = "1.0.0"
    enabled = True

    extra_columns = ['po_number', 'unit_price', 'description', 'country_origin']

    # Keywords to identify this supplier
    SUPPLIER_KEYWORDS = [
        'Invoice No'
    ]

    def __init__(self):
        super().__init__()
        self._extractor = None
        self._last_result = None

    @property
    def extractor(self):
        """Lazy-load SmartExtractor."""
        if self._extractor is None and SmartExtractor is not None:
            self._extractor = SmartExtractor()
        return self._extractor

    def can_process(self, text: str) -> bool:
        """Check if this is a Sigma_Shaanxi invoice."""
        text_lower = text.lower()

        # Check for supplier keywords
        for keyword in self.SUPPLIER_KEYWORDS:
            if keyword in text_lower:
                return True

        return False

    def get_confidence_score(self, text: str) -> float:
        """Return confidence score for template matching."""
        if not self.can_process(text):
            return 0.0

        score = 0.7  # High base score for specific supplier match
        text_lower = text.lower()

        # Add confidence for each keyword found
        for keyword in self.SUPPLIER_KEYWORDS:
            if keyword in text_lower:
                score += 0.1
                break

        # Add confidence for client name
        if 'sigma corporation' in text_lower:
            score += 0.1

        return min(score, 1.0)

    def extract_invoice_number(self, text: str) -> str:
        """Extract invoice number."""
        patterns = [
            r'INVOICE\s*(?:NO\.?)?\s*[:\s]*([A-Z0-9][\w\-/]+)',
            r'Invoice\s*(?:No\.?|#)\s*[:\s]*([A-Z0-9][\w\-/]+)',
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()

        return "UNKNOWN"

    def extract_project_number(self, text: str) -> str:
        """Extract PO number."""
        patterns = [
            r'P\.?O\.?\s*#?\s*:?\s*(\d{6,})',
            r'Purchase\s*Order[:\s]*(\d+)',
            r'\b(400\d{5})\b',  # Sigma PO format
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)

        return "UNKNOWN"

    def extract_manufacturer_name(self, text: str) -> str:
        """Return the manufacturer name."""
        return "SIGMA_SHAANXI"

    def extract_line_items(self, text: str) -> List[Dict]:
        """
        Extract line items using SmartExtractor.
        """
        if not self.extractor:
            return []

        try:
            self._last_result = self.extractor.extract_from_text(text)

            items = []
            for item in self._last_result.line_items:
                items.append({
                    'part_number': item.part_number,
                    'quantity': item.quantity,
                    'total_price': item.total_price,
                    'unit_price': item.unit_price,
                    'description': item.description,
                    'po_number': self._last_result.po_numbers[0] if self._last_result.po_numbers else '',
                    'country_origin': 'CHINA',
                })

            return items

        except Exception as e:
            print(f"SmartExtractor error: {e}")
            return []

    def post_process_items(self, items: List[Dict]) -> List[Dict]:
        """Post-process - deduplicate."""
        if not items:
            return items

        seen = set()
        unique_items = []

        for item in items:
            key = f"{item['part_number']}_{item['quantity']}_{item['total_price']}"
            if key not in seen:
                seen.add(key)
                unique_items.append(item)

        return unique_items

    def is_packing_list(self, text: str) -> bool:
        """Check if document is only a packing list."""
        text_lower = text.lower()
        if 'packing list' in text_lower and 'invoice' not in text_lower:
            return True
        return False
