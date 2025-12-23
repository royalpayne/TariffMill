"""
SigmaSuppliersTemplate - Invoice template for Sigma Corporation suppliers
Handles invoices from multiple suppliers including:
- PT. Laju Sinergi Metalindo (Indonesia)
- Crescent Foundry Co Pvt. Ltd (India)
- King Multimetal Industries Private Limited (India)
- Himgiri Castings Pvt. Ltd (India)
- Calcutta Springs Limited (India)
- Seksaria Foundries Limited (India)
- Sri Ranganathar Industries (India)
"""

import re
from typing import List, Dict
from .base_template import BaseTemplate


class SigmaSuppliersTemplate(BaseTemplate):
    """
    Template for Sigma Corporation supplier invoices.
    Handles various Indian and Indonesian manufacturers.
    """

    name = "Sigma Suppliers"
    description = "Invoices from Sigma Corporation suppliers (India, Indonesia)"
    client = "Sigma Corporation"
    version = "1.0.0"
    enabled = True

    extra_columns = ['po_number', 'hs_code', 'unit_price', 'description', 'country_origin']

    # Known suppliers
    SUPPLIERS = [
        'laju sinergi metalindo',
        'crescent foundry',
        'king multimetal',
        'himgiri castings',
        'calcutta springs',
        'seksaria foundries',
        'sri ranganathar',
        'rba exports',
    ]

    def can_process(self, text: str) -> bool:
        """Check if this template can process the invoice."""
        text_lower = text.lower()

        # Must be for Sigma Corporation
        if 'sigma corporation' not in text_lower and 'sigma corp' not in text_lower:
            return False

        # Must be from one of the known suppliers
        for supplier in self.SUPPLIERS:
            if supplier in text_lower:
                return True

        return False

    def get_confidence_score(self, text: str) -> float:
        """Return confidence score for template matching."""
        if not self.can_process(text):
            return 0.0

        score = 0.5
        text_lower = text.lower()

        # Add confidence for each supplier found
        for supplier in self.SUPPLIERS:
            if supplier in text_lower:
                score += 0.2
                break

        # Add confidence for invoice markers
        if re.search(r'invoice\s*(?:no\.?|#)\s*:?\s*[\w\-/]+', text, re.IGNORECASE):
            score += 0.15
        if re.search(r'po\.?\s*(?:no\.?|#)\s*:?\s*\d+', text, re.IGNORECASE):
            score += 0.1
        if 'india' in text_lower or 'indonesia' in text_lower:
            score += 0.1

        return min(score, 1.0)

    def extract_invoice_number(self, text: str) -> str:
        """Extract invoice number."""
        patterns = [
            r'INVOICE\s*(?:NO\.?|#)\s*[&:]?\s*(?:DATE:?)?\s*([\w\-/]+)',  # INVOICE NO: xxx or INVOICE NO. & DATE: xxx
            r'Invoice\s+No\.?\s*:?\s*([\w\-/]+)',
            r'Invoice\s+(?:Number|#)\s*:?\s*([\w\-/]+)',
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                inv_num = match.group(1).strip()
                # Clean up if it starts with 'DATE' or contains date
                if inv_num.upper().startswith('DATE'):
                    continue
                return inv_num

        return "UNKNOWN"

    def extract_project_number(self, text: str) -> str:
        """Extract PO number as project reference."""
        patterns = [
            r'PO\.?\s*NO\.?\s*[:\s]*(\d{8})',                    # PO. NO. 40049433
            r'Buyer[\'s]*\s*[Oo]rder\s*[Nn]o\.?\s*:?\s*(\d+)',  # Buyer order No: 40050092
            r'Order\s*No\s*:?\s*(\d+)',                          # Order No: xxx
            r'\b(400\d{5})\b',                                   # Sigma PO format: 400xxxxx
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()

        return "UNKNOWN"

    def extract_manufacturer_name(self, text: str) -> str:
        """Extract manufacturer name."""
        text_lower = text.lower()

        supplier_map = {
            'laju sinergi metalindo': 'PT. LAJU SINERGI METALINDO',
            'crescent foundry': 'CRESCENT FOUNDRY CO PVT. LTD',
            'king multimetal': 'KING MULTIMETAL INDUSTRIES PRIVATE LIMITED',
            'himgiri castings': 'HIMGIRI CASTINGS PVT. LTD',
            'calcutta springs': 'CALCUTTA SPRINGS LIMITED',
            'seksaria foundries': 'SEKSARIA FOUNDRIES LIMITED',
            'sri ranganathar': 'SRI RANGANATHAR INDUSTRIES (P) LTD',
            'rba exports': 'RBA EXPORTS PRIVATE LIMITED',
        }

        for key, name in supplier_map.items():
            if key in text_lower:
                return name

        return ""

    def extract_line_items(self, text: str) -> List[Dict]:
        """Extract line items from invoice."""
        line_items = []
        seen_items = set()

        # Determine invoice type based on content
        text_lower = text.lower()

        if 'laju sinergi' in text_lower:
            line_items = self._extract_laju_items(text)
        elif 'crescent foundry' in text_lower:
            line_items = self._extract_crescent_items(text)
        elif 'king multimetal' in text_lower:
            line_items = self._extract_king_multimetal_items(text)
        elif 'himgiri castings' in text_lower:
            line_items = self._extract_himgiri_items(text)
        elif 'calcutta springs' in text_lower:
            line_items = self._extract_calcutta_springs_items(text)
        elif 'seksaria foundries' in text_lower:
            line_items = self._extract_seksaria_items(text)
        elif 'sri ranganathar' in text_lower:
            line_items = self._extract_sri_ranganathar_items(text)
        else:
            # Generic extraction
            line_items = self._extract_generic_items(text)

        # Deduplicate
        unique_items = []
        for item in line_items:
            item_key = f"{item['part_number']}_{item['quantity']}_{item['total_price']}"
            if item_key not in seen_items:
                seen_items.add(item_key)
                unique_items.append(item)

        return unique_items

    def _extract_laju_items(self, text: str) -> List[Dict]:
        """Extract items from Laju Sinergi Metalindo invoices."""
        items = []
        lines = text.split('\n')

        # Pattern: Description Size Item Code Ctn Pcs/Ctn Pcs Price/Pcs Total
        # Example: - HUB 274301 DI FITTING STRAIN EYE - 18-2743014001 8 2,000 16,000 1.63 26,080.00
        pattern = re.compile(
            r'[-\s]*HUB\s+(\S+)\s+'                              # HUB + part identifier
            r'(.+?)\s+'                                          # Description
            r'(18-[\w]+)\s+'                                     # Item Code (18-xxx format)
            r'(\d+)\s+'                                          # Ctn count
            r'([\d,]+)\s+'                                       # Pcs/Ctn
            r'([\d,]+)\s+'                                       # Total Pcs
            r'([\d.]+)\s+'                                       # Price/Pcs
            r'([\d,.]+)',                                        # Total
            re.IGNORECASE
        )

        # Also try simpler pattern for single-line items
        simple_pattern = re.compile(
            r'(18-[\w]+)\s+'                                     # Item Code
            r'(\d+)\s+'                                          # Ctn
            r'([\d,]+)\s+'                                       # Pcs/Ctn
            r'([\d,]+)\s+'                                       # Total Pcs
            r'([\d.]+)\s+'                                       # Price/Pcs
            r'([\d,.]+)',                                        # Total
            re.IGNORECASE
        )

        # Extract PO number from context
        po_match = re.search(r'PO\.?\s*NO\.?\s*(\d+)', text, re.IGNORECASE)
        current_po = po_match.group(1) if po_match else ''

        for match in simple_pattern.finditer(text):
            try:
                item_code = match.group(1)
                total_pcs = match.group(4).replace(',', '')
                unit_price = match.group(5)
                total_price = match.group(6).replace(',', '')

                items.append({
                    'part_number': item_code,
                    'quantity': total_pcs,
                    'total_price': total_price,
                    'unit_price': unit_price,
                    'description': '',
                    'po_number': current_po,
                    'hs_code': '7326.90.86',
                    'country_origin': 'INDONESIA',
                })
            except (ValueError, IndexError):
                continue

        return items

    def _extract_crescent_items(self, text: str) -> List[Dict]:
        """Extract items from Crescent Foundry invoices."""
        items = []

        # Pattern: ITEM CODE | DESCRIPTION | ORDER NO | ORDER DT | HSN CODE | QTY | UNIT | RATE | AMOUNT
        # Example: NMS-V-004 | COUNTERWEIGHT(HP 10500) | 40049125 | 20250918 | 84312010 | PC | USD 2676.00 | USD 13380.00
        pattern = re.compile(
            r'([\w\-]+)\s+'                                      # Item code
            r'(?:Parts of [\w\s]+\(of machinery of heading \d+\s*\))\s*'  # Description (parts of...)
            r'(\d{8})\s+'                                        # Order No (PO)
            r'\d{8}\s+'                                          # Order Date
            r'(\d{8})\s+'                                        # HSN Code
            r'(\w+)\s+'                                          # Unit
            r'USD\s*([\d,.]+)\s+'                                # Rate
            r'USD\s*([\d,.]+)',                                  # Amount
            re.IGNORECASE
        )

        for match in pattern.finditer(text):
            try:
                items.append({
                    'part_number': match.group(1),
                    'quantity': '1',  # Usually bulk items
                    'total_price': match.group(6).replace(',', ''),
                    'unit_price': match.group(5).replace(',', ''),
                    'description': 'Parts of fork lift trucks',
                    'po_number': match.group(2),
                    'hs_code': match.group(3),
                    'country_origin': 'INDIA',
                })
            except (ValueError, IndexError):
                continue

        # Simpler fallback pattern
        if not items:
            fallback = re.compile(
                r'(NMS-[\w\-]+)\s+'                              # Item code NMS-xxx
                r'.*?(\d{8})\s+'                                 # PO number
                r'.*?(\d{8})\s+'                                 # Date
                r'(\d{8})\s+'                                    # HSN
                r'(\w+)\s+'                                      # Unit
                r'USD\s*([\d,.]+)\s+'                            # Rate
                r'USD\s*([\d,.]+)',                              # Amount
                re.IGNORECASE | re.DOTALL
            )
            for match in fallback.finditer(text):
                try:
                    # Extract quantity from nearby context
                    qty_match = re.search(r'(\d+)\s*' + re.escape(match.group(5)), text)
                    qty = qty_match.group(1) if qty_match else '1'

                    items.append({
                        'part_number': match.group(1),
                        'quantity': qty,
                        'total_price': match.group(7).replace(',', ''),
                        'unit_price': match.group(6).replace(',', ''),
                        'description': '',
                        'po_number': match.group(2),
                        'hs_code': match.group(4),
                        'country_origin': 'INDIA',
                    })
                except (ValueError, IndexError):
                    continue

        return items

    def _extract_king_multimetal_items(self, text: str) -> List[Dict]:
        """Extract items from King Multimetal invoices."""
        items = []

        # Pattern: Description | HSN CODE | PO NO | ITEM CODE | QTY IN PCS | RATE | AMOUNT
        # Example: HUBBELL 681812 WASHER... | 73182200 | 40050092 | 18-681812 | 18800 | 0.7150 | 13,442.00
        pattern = re.compile(
            r'(\d+)\s+'                                          # Line number
            r'(\d{8})\s+'                                        # HSN Code
            r'(\d{8})\s+'                                        # PO Number
            r'(18-[\w]+)\s+'                                     # Item Code
            r'(\d+)\s+'                                          # Qty
            r'([\d.]+)\s+'                                       # Rate
            r'([\d,.]+)',                                        # Amount
            re.IGNORECASE
        )

        for match in pattern.finditer(text):
            try:
                items.append({
                    'part_number': match.group(4),
                    'quantity': match.group(5),
                    'total_price': match.group(7).replace(',', ''),
                    'unit_price': match.group(6),
                    'description': '',
                    'po_number': match.group(3),
                    'hs_code': match.group(2),
                    'country_origin': 'INDIA',
                })
            except (ValueError, IndexError):
                continue

        return items

    def _extract_himgiri_items(self, text: str) -> List[Dict]:
        """Extract items from Himgiri Castings invoices."""
        items = []
        # Similar structure to Crescent - implement as needed
        return items

    def _extract_calcutta_springs_items(self, text: str) -> List[Dict]:
        """Extract items from Calcutta Springs invoices."""
        items = []
        # CSL format - implement as needed
        return items

    def _extract_seksaria_items(self, text: str) -> List[Dict]:
        """Extract items from Seksaria Foundries invoices."""
        items = []
        # Seksaria format - implement as needed
        return items

    def _extract_sri_ranganathar_items(self, text: str) -> List[Dict]:
        """Extract items from Sri Ranganathar invoices."""
        items = []
        # SRI format - implement as needed
        return items

    def _extract_generic_items(self, text: str) -> List[Dict]:
        """Generic extraction for unknown formats."""
        items = []

        # Try to find item code + quantity + price patterns
        pattern = re.compile(
            r'\b(18-[\w]+|\w{2,4}-[\w\-]+)\b\s+'                # Item code
            r'.*?(\d+(?:,\d{3})*)\s+'                           # Quantity
            r'.*?\$?\s*([\d,]+\.?\d*)',                         # Price
            re.IGNORECASE
        )

        for match in pattern.finditer(text):
            try:
                items.append({
                    'part_number': match.group(1),
                    'quantity': match.group(2).replace(',', ''),
                    'total_price': match.group(3).replace(',', ''),
                    'unit_price': '',
                    'description': '',
                    'po_number': '',
                    'hs_code': '',
                    'country_origin': '',
                })
            except (ValueError, IndexError):
                continue

        return items

    def is_packing_list(self, text: str) -> bool:
        """Check if document is only a packing list."""
        text_lower = text.lower()
        if 'packing list' in text_lower:
            # If it also contains invoice, process it
            if 'invoice' in text_lower:
                return False
            return True
        return False
