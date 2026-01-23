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
            # Patrón original con más contexto
            r'[Dd]irecci[oó]n[:\s]+(.+?)(?:\n|Referencia|CP|C\.P\.|Catastral)',
            r'[Dd]omicilio[:\s]+(.+?)(?:\n|CP|C\.P\.|Referencia)',
            
            # Patrón más flexible
            r'[Dd]irecci[oó]n\s*:\s*([^\n]{10,150})',
            r'[Dd]omicilio\s*:\s*([^\n]{10,150})',
            
            # Patrón genérico: busca direcciones tipo "CL ... CP ..."
            r'((?:CL|CALLE|AV|AVENIDA|PLAZA|PL)\s+[A-ZÁÉÍÓÚÑ\s,\d]{10,100})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE | re.DOTALL)
            if match:
                addr = match.group(1).strip()
                # Limpia saltos de línea
                addr = re.sub(r'\s+', ' ', addr)
                if len(addr) >= 10:
                    return addr
        
        return "NOT FOUND"

    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference"""
        # 1) Patrón con contexto "referencia catastral"
        pattern = r'referencia\s+catastral[^0-9A-Z]{0,30}([0-9A-Z\s]{14,25})'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            catastral = match.group(1).strip()
            catastral = re.sub(r'\s+', '', catastral)  # quita espacios
            if re.match(r'^[0-9]+[A-Z]+[0-9]+[A-Z]+[0-9]+[A-Z]+$', catastral):
                return catastral
        
        # 2) Patrón directo: busca formato típico 14-20 chars (números + letras)
        match = re.search(r'\b(\d{6,8}[A-Z]{1,3}\d{2,6}[A-Z]{1,4}\d{1,6}[A-Z]{1,4})\b', self.text)
        if match:
            return match.group(1).upper()
        
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
