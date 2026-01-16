"""
Parser for FACTURA documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class FacturaParser(BaseDocumentParser):
    """Parser for Factura"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Factura"""
        self.extract_text()
        
        result = {
            'document_type': 'FACTURA',
            'invoice_number': self._extract_invoice_number(),
            'invoice_date': self._extract_invoice_date(),
            'homeowner_name': self._extract_homeowner_name(),
            'homeowner_dni': self._extract_homeowner_dni(),
            'homeowner_address': self._extract_homeowner_address(),
            'amount': self._extract_amount(),
        }
        
        return result
    
    def _extract_invoice_number(self) -> str:
        """Extract invoice number"""
        patterns = [
            r'[Ff]actura[:\s]+n[úuº.]*\s*([A-Z0-9-]+)',
            r'[Nn][úuº.]*\s*[Ff]actura[:\s]+([A-Z0-9-]+)',
            r'N[úº]\s+([A-Z0-9-]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_invoice_date(self) -> str:
        """Extract invoice date"""
        patterns = [
            r'[Ff]echa[:\s]+([\d/]+)',
            r'(\d{2}/\d{2}/\d{4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_homeowner_name(self) -> str:
        """Extract customer/homeowner name"""
        patterns = [
            r'[Cc]liente[:\s]+([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+(?:\s+[A-ZÑ][a-zñ]+)?)',
            r'[Nn]ombre[:\s]+([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_homeowner_dni(self) -> str:
        """Extract DNI"""
        match = re.search(r'\b(\d{8}[A-Z])\b', self.text)
        if match:
            return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_homeowner_address(self) -> str:
        """Extract address"""
        patterns = [
            r'[Dd]irecci[oó]n[:\s]+(.+?)(?:\n|CP|$)',
            r'[Dd]omicilio[:\s]+(.+?)(?:\n|CP|$)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_amount(self) -> str:
        """Extract total amount"""
        patterns = [
            r'[Tt]otal[:\s]+([\d.,]+)\s*€',
            r'[Ii]mporte[:\s]+([\d.,]+)\s*€',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return f"{match.group(1)} €"
        
        return "NOT FOUND"
