"""
Parser for DECLARACION RESPONSABLE documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class DeclaracionParser(BaseDocumentParser):
    """Parser for Declaración Responsable"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Declaracion"""
        self.extract_text()
        
        result = {
            'document_type': 'DECLARACION_RESPONSABLE',
            'homeowner_name': self._extract_homeowner_name(),
            'homeowner_dni': self._extract_homeowner_dni(),
            'homeowner_address': self._extract_homeowner_address(),
            'catastral_ref': self._extract_catastral_ref(),
            'act_code': self._extract_act_code(),
            'signature': self._check_signature(),
        }
        
        return result
    
    def _extract_homeowner_name(self) -> str:
        """Extract homeowner name"""
        patterns = [
            r'D[./]?\s*([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+(?:\s+[A-ZÑ][a-zñ]+)?)',
            r'titular[:\s]+([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+)',
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
            r'[Dd]omicilio[:\s]+(.+?)(?:\n|,\s*\d{5})',
            r'[Dd]irecci[oó]n[:\s]+(.+?)(?:\n|CP)',
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
    
    def _extract_act_code(self) -> str:
        """Extract RES code"""
        match = re.search(r'(RES0*\d{2,3})', self.text, re.IGNORECASE)
        if match:
            code = match.group(1).upper()
            code = re.sub(r'RES0+(\d)', r'RES0\1', code)
            return code
        
        return "NOT FOUND"
    
    def _check_signature(self) -> str:
        """Check if signature exists"""
        if re.search(r'Firma|Firmado|Fdo\.', self.text, re.IGNORECASE):
            return "Present"
        
        return "NOT FOUND"
