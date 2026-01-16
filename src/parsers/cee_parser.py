"""
Parser for CEE FINAL documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class CeeParser(BaseDocumentParser):
    """Parser for CEE Final"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from CEE"""
        self.extract_text()
        
        result = {
            'document_type': 'CEE_FINAL',
            'address': self._extract_address(),
            'catastral_ref': self._extract_catastral_ref(),
            'certification_date': self._extract_certification_date(),
            'signature': self._check_signature(),
        }
        
        return result
    
    def _extract_address(self) -> str:
        """Extract property address"""
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
    
    def _extract_certification_date(self) -> str:
        """Extract certification date"""
        patterns = [
            r'[Ff]echa[:\s]+([\d/]+)',
            r'(\d{2}/\d{2}/\d{4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _check_signature(self) -> str:
        """Check if signature exists"""
        if re.search(r'Firma|Firmado|Fdo\.', self.text, re.IGNORECASE):
            return "Present"
        
        return "NOT FOUND"
