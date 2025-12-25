"""
Shaanxi Fangzhi Trade Co Template

Auto-generated template for invoices from Shaanxi Fangzhi Trade Co.
Generated: 2025-12-24 12:56:59
"""

import re
from typing import List, Dict
from .base_template import BaseTemplate


class ShaanxiFangzhiTradeTemplate(BaseTemplate):
    """Template for Shaanxi Fangzhi Trade Co invoices."""

    name = "Shaanxi Fangzhi Trade Co"
    description = "Invoices from Shaanxi Fangzhi Trade Co"
    client = "Sigma Corporation"
    version = "1.0.0"
    enabled = True

    extra_columns = ['po_number', 'unit_price', 'description', 'country_origin']

    SUPPLIER_KEYWORDS = [
        r'shaanxi fangzhi trade co',
        r'\bshanzhou\b'  # Add Shanzhou keyword here
    ]

    def can_process(self, text: str) -> bool:
        """Check if this is a Shaanxi Fangzhi Trade Co invoice."""
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
        pattern = r'invoice no\s*:\s*([a-zA-Z0-9]+)'
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return "UNKNOWN"

    def extract_project_number(self, text: str) -> str:
        """Extract PO/project number."""
        patterns = [
            r'p\.\s*\d{4,6}/\d+',
            # Add patterns for PO numbers
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return "UNKNOWN"

    def extract_manufacturer_name(self, text: str) -> str:
        """Return the manufacturer name."""
        return "SHAANXI FANGZHI TRADE CO"

    def extract_line_items(self, text: str) -> List[Dict]:
        """Extract line items from invoice."""
        patterns = [
            r'\s*([^;]+)\s+([0-9.]+)\s+([0-9.]+)',  # part_number quantity price
        ]
        for pattern in patterns:
            match = re.findall(pattern, text, re.IGNORECASE)
            if match:
                items = []
                for item in match:
                    items.append({
                        'part_number': item[0],
                        'quantity': item[1].strip(),
                        'total_price': item[2].strip(),
                    })
        return items

    def post_process_items(self, items: List[Dict]) -> List[Dict]:
        """Post-process - deduplicate and validate."""
        if not items:
            return items

        seen = set()
        unique_items = []

        for item in items:
            key = f"{item['part_number']}_{item['quantity']}_{item['total_price']}"
            if key not in seen:
                seen.add(key)
                # Add country of origin
                item['country_origin'] = 'CHINA'
                unique_items.append(item)

        return unique_items

    def is_packing_list(self, text: str) -> bool:
        """Check if document is only a packing list."""
        text_lower = text.lower()
        if 'packing list' in text_lower and 'invoice' not in text_lower:
            return True
        return False