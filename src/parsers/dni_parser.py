"""
Parser for DNI documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class DniParser(BaseDocumentParser):
    """Parser for DNI"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from DNI"""
        self.extract_text()
        
        result = {
            'document_type': 'DNI',
            'dni_number': self._extract_dni_number(),
            'name': self._extract_name(),
        }
        
        return result
    
    def _extract_dni_number(self) -> str:
        """Extract DNI number"""
        match = re.search(r'\b(\d{8}[A-Z])\b', self.text)
        if match:
            return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_name(self) -> str:
        """Extract name from DNI"""
        patterns = [
            r'[Nn]ombre[:\s]+([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+)',
            r'Apellidos[:\s]+(.+?)(?:\n|Nombre)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
