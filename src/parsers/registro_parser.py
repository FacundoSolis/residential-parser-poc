"""
Parser for REGISTRO CEE documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class RegistroParser(BaseDocumentParser):
    """Parser for Registro CEE"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Registro"""
        self.extract_text()
        
        result = {
            'document_type': 'REGISTRO_CEE',
            'registration_date': self._extract_registration_date(),
            'registration_number': self._extract_registration_number(),
            'address': self._extract_address(),
            'catastral_ref': self._extract_catastral_ref(),
        }
        
        return result
    
    def _extract_registration_date(self) -> str:
        """Extract registration date"""
        patterns = [
            r'[Ff]echa[:\s]+(?:de\s+)?(?:registro|inscripci[oó]n)[:\s]+([\d/]+)',
            r'(\d{2}/\d{2}/\d{4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_registration_number(self) -> str:
        """Extract registration number"""
        patterns = [
            r'[Nn][úuº.]*\s*[Rr]egistro[:\s]+([A-Z0-9-]+)',
            r'[Rr]egistro[:\s]+n[úuº.]*\s*([A-Z0-9-]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_address(self) -> str:
        """Extract address"""
        patterns = [
            r'[Dd]irecci[oó]n[:\s]+(.+?)(?:\n|Referencia)',
            r'[Dd]omicilio[:\s]+(.+?)(?:\n|CP)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference"""
        pattern = r'referencia catastral[^0-9A-Z]*([0-9A-Z]+)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            catastral = match.group(1).strip()
            catastral = re.sub(r'\s+', '', catastral)
            if re.match(r'^[0-9]+[A-Z]+[0-9]+[A-Z]+[0-9]+[A-Z]+$', catastral):
                return catastral
        
        return "NOT FOUND"
