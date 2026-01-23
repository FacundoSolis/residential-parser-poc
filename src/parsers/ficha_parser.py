"""
Parser for FICHA RES020 documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class FichaParser(BaseDocumentParser):
    """Parser for Ficha RES020 documents"""

    def _normalize(self, t: str) -> str:
        """Normalize text for better pattern matching"""
        if not t:
            return ""

        # Basic normalization for OCR text
        t = (
            t.replace("’", "'")
            .replace("´", "'")
            .replace("`", "'")
            .replace("″", "''")
        )

        # Common OCR errors in Spanish
        t = t.replace("Le6n", "León").replace("Direcci6n", "Dirección").replace("ubicaci6n", "ubicación")

        return t

    def parse(self) -> Dict[str, Any]:
        self.extract_text()
        t = self._normalize(self.text)

        return {
            'document_type': 'FICHA',
            'homeowner_name': self._extract_homeowner_name(t),
            'homeowner_address': self._extract_homeowner_address(t),
            'homeowner_dni': self._extract_homeowner_dni(t),
            'act_code': self._extract_act_code(t),
            'catastral_ref': self._extract_catastral_ref(t),
            'energy_savings': self._extract_energy_savings(t),
            'fp': self._extract_fp(t),
            'ui': self._extract_ui(t),
            'uf': self._extract_uf(t),
            'surface': self._extract_surface(t),
            'climatic_zone': self._extract_climatic_zone(t),
            'isolation_thickness': self._extract_isolation_thickness(t),
        }

    def _extract_homeowner_name(self, text: str) -> str:
        """Extract homeowner name from ficha"""
        patterns = [
            r'(?:propietario|titular|dueño)[:\s]*([^\n\r]{10,50})',
            r'(?:nombre|name)[:\s]*([^\n\r]{10,50})',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_homeowner_address(self, text: str) -> str:
        """Extract homeowner address from ficha"""
        patterns = [
            r'(?:domicilio|dirección|address)[:\s]*([^\n\r]{15,80})',
            r'(?:ubicación|location)[:\s]*([^\n\r]{15,80})',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_homeowner_dni(self, text: str) -> str:
        """Extract homeowner DNI from ficha"""
        patterns = [
            r'(?:dni|nie|documento)[:\s]*([0-9]{8}[A-Z]|[XYZ][0-9]{7}[A-Z])',
            r'(?:nif|cif)[:\s]*([0-9]{8}[A-Z]|[XYZ][0-9]{7}[A-Z])',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_act_code(self, text: str) -> str:
        """Extract ACT code from ficha - busca RES010 o RES020"""
        m = re.search(r"(RES0*\d{2,3})", text, re.IGNORECASE)
        if m:
            code = m.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)  # normaliza RES00020 -> RES020
            return code
        
        patterns = [
            r'(?:código\s*act|act\s*code)[:\s]*(RES\s*0*\d{2,3})',
            r'(?:tipo\s*de\s*act|act\s*type)[:\s]*(RES\s*0*\d{2,3})',
            r'(?:medida|actuación)[:\s]*(RES\s*0*\d{2,3})',
        ]
        
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                code = m.group(1).upper().replace(" ", "")
                code = re.sub(r"RES0+(\d)", r"RES0\1", code)
                return code
        
        return ""

    def _extract_catastral_ref(self, text: str) -> str:
        """Extract catastral reference from ficha"""
        patterns = [
            r'(?:referencia\s*catastral|catastral\s*ref)[:\s]*([0-9A-Z\s]{14,20})',
            r'(?:ref\s*catastral)[:\s]*([0-9A-Z\s]{14,20})',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_energy_savings(self, text: str) -> str:
        """Extract energy savings from ficha"""
        patterns = [
            r'(?:ahorro\s*energético|energy\s*savings)[:\s]*([0-9.,]+)',
            r'(?:kwh|kilowatts? hora)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_fp(self, text: str) -> str:
        """Extract Fp from ficha"""
        patterns = [
            r'(?:fp|factor\s*p)[:\s]*([0-9.,]+)',
            r'(?:p\s*factor)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_ui(self, text: str) -> str:
        """Extract Ui from ficha"""
        patterns = [
            r'(?:ui|transmitancia)[:\s]*([0-9.,]+)',
            r'(?:u\s*inicial)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_uf(self, text: str) -> str:
        """Extract Uf from ficha"""
        patterns = [
            r'(?:uf|u\s*final)[:\s]*([0-9.,]+)',
            r'(?:transmitancia\s*final)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_surface(self, text: str) -> str:
        """Extract surface area from ficha"""
        patterns = [
            r'(?:superficie|surface|área)[:\s]*([0-9.,]+)',
            r'(?:s\s*=|superficie\s*=)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_climatic_zone(self, text: str) -> str:
        """Extract climatic zone from ficha"""
        patterns = [
            r'(?:zona\s*climática|climatic\s*zone)[:\s]*([A-Z0-9]+)',
            r'(?:zona\s*=|zone\s*=)[:\s]*([A-Z0-9]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_isolation_thickness(self, text: str) -> str:
        """Extract isolation thickness from ficha"""
        patterns = [
            r'(?:espesor\s*aislamiento|isolation\s*thickness)[:\s]*([0-9.,]+)',
            r'(?:grosor\s*aislante)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _find_first_pattern(self, text: str, patterns: list) -> str:
        """Find first matching pattern"""
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                result = match.group(1).strip()
                if len(result) > 3:  # Avoid garbage matches
                    return result
        return ""
