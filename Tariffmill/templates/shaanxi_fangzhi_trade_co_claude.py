"""
Shaanxi Fangzhi Trade Co Template

Auto-generated template for invoices from Shaanxi Fangzhi Trade Co.
Generated: 2025-12-24 13:56:09
"""

import re
from typing import List, Dict
from .base_template import BaseTemplate


class ShaanxiFangzhiTradeCoClaudeTemplate(BaseTemplate):
    """Template for Shaanxi Fangzhi Trade Co invoices."""

    name = "Shaanxi Fangzhi Trade Co"
    description = "Invoices from Shaanxi Fangzhi Trade Co"
    client = "Sigma Corporation"
    version = "1.0.0"
    enabled = True

    extra_columns = ['po_number', 'unit_price', 'description', 'country_origin']

    # Keywords to identify this supplier
    SUPPLIER_KEYWORDS = [
        "shaanxi fangzhi trade co",
        "green mansion no.1 of xingqing road middle",
        "xi'an, china",
        "messer.sigma corporation",
        "cream ridge,nj08514",
        "onelok(pipe restraint products)",
        "pv lok/restrainer"
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
        patterns = [
            r"INVOICE NO\.([A-Z0-9]+)",
            r"INVOICE\s+NO\.\s*([A-Z0-9]+)",
            r"INVOICE\s+NO[:\.]?\s*([A-Z0-9]+)"
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return "UNKNOWN"

    def extract_project_number(self, text: str) -> str:
        """Extract PO/project number."""
        patterns = [
            r"P\.O\.#:\s*(\d+)",
            r"P\.O\.\s*#:\s*(\d+)",
            r"P\.O\.#\s*(\d+)",
            r"CONTRACT NO\s*(\d+)"
        ]
        po_numbers = []
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            po_numbers.extend(matches)
        
        if po_numbers:
            # Return the first PO number found
            return po_numbers[0]
        return "UNKNOWN"

    def extract_manufacturer_name(self, text: str) -> str:
        """Return the manufacturer name."""
        return "SHAANXI FANGZHI TRADE CO"

    def extract_line_items(self, text: str) -> List[Dict]:
        """Extract line items from invoice."""
        items = []
        
        # Pattern for simple line items (part_number part_number2 qty qty unit_price total_price)
        simple_pattern = r"([A-Z0-9\-]+)\s+([A-Z0-9\-]+)\s+(\d+[\,\d]*)\s+(\d+[\,\d]*)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)"
        
        # Pattern for complex line items with accessories (SLDEP series)
        complex_pattern = r"(SLDEP\d+[A-Z0-9\-\+\/]+)\s+(SLDEP\d+)\s+(\d+[\,\d]*)\s+(\d+[\,\d]*)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)\s+(\d+[\,\d]*)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)\s+(\d+[\,\d]*)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)\s+(\d+[\,\d]*)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)"
        
        # Extract current PO number context
        current_po = "UNKNOWN"
        po_pattern = r"P\.O\.#:\s*(\d+)"
        po_matches = re.findall(po_pattern, text)
        
        lines = text.split('\n')
        current_po = "UNKNOWN"
        
        for i, line in enumerate(lines):
            # Track current PO number
            po_match = re.search(po_pattern, line)
            if po_match:
                current_po = po_match.group(1)
                continue
            
            # Try simple pattern first
            simple_match = re.search(simple_pattern, line)
            if simple_match:
                part_number = simple_match.group(1)
                alt_part_number = simple_match.group(2)
                quantity = int(simple_match.group(3).replace(',', ''))
                unit_price = float(simple_match.group(5).replace(',', ''))
                total_price = float(simple_match.group(6).replace(',', ''))
                
                items.append({
                    'part_number': part_number,
                    'quantity': quantity,
                    'unit_price': unit_price,
                    'total_price': total_price,
                    'po_number': current_po,
                    'description': alt_part_number,
                    'country_origin': 'CHINA'
                })
                continue
            
            # Try complex pattern for SLDEP items
            complex_match = re.search(complex_pattern, line)
            if complex_match:
                part_number = complex_match.group(1)
                alt_part_number = complex_match.group(2)
                quantity = int(complex_match.group(3).replace(',', ''))
                unit_price = float(complex_match.group(5).replace(',', ''))
                total_price = float(complex_match.group(16).replace(',', ''))
                
                items.append({
                    'part_number': part_number,
                    'quantity': quantity,
                    'unit_price': unit_price,
                    'total_price': total_price,
                    'po_number': current_po,
                    'description': alt_part_number,
                    'country_origin': 'CHINA'
                })
                continue
        
        # Fallback: Extract any part numbers that look like product codes
        if not items:
            fallback_pattern = r"([A-Z]{2,}[A-Z0-9\-]+)\s+([A-Z]{2,}[A-Z0-9\-]*)\s+([0-9,]+)\s+([0-9,]+)\s+\$([0-9,]+\.?\d*)\s+\$([0-9,]+\.?\d*)"
            fallback_matches = re.findall(fallback_pattern, text)
            
            for match in fallback_matches:
                try:
                    part_number = match[0]
                    description = match[1]
                    quantity = int(match[2].replace(',', ''))
                    unit_price = float(match[4].replace(',', ''))
                    total_price = float(match[5].replace(',', ''))
                    
                    items.append({
                        'part_number': part_number,
                        'quantity': quantity,
                        'unit_price': unit_price,
                        'total_price': total_price,
                        'po_number': current_po,
                        'description': description,
                        'country_origin': 'CHINA'
                    })
                except (ValueError, IndexError):
                    continue

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