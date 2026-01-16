"""
Parser for CONTRATO CESION AHORROS documents
Handles encoding issues in PDFs (6 instead of ó, f instead of í)
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class ContratoParser(BaseDocumentParser):
    """Parser for Contrato Cesión Ahorros"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Contrato"""
        self.extract_text()
        
        result = {
            'document_type': 'CONTRATO_CESION_AHORROS',
            'sd_do_company_name': self._extract_cesionario_company(),
            'sd_do_cif': self._extract_cesionario_cif(),
            'sd_do_address': self._extract_cesionario_address(),
            'sd_do_representative_name': self._extract_cesionario_representative(),
            'sd_do_representative_dni': self._extract_cesionario_dni(),
            'sd_do_signature': self._check_cesionario_signature(),
            'homeowner_name': self._extract_cedente_name(),
            'homeowner_dni': self._extract_cedente_dni(),
            'homeowner_address': self._extract_cedente_address(),
            'homeowner_phone': self._extract_cedente_phone(),
            'homeowner_email': self._extract_cedente_email(),
            'homeowner_notifications': self._extract_notifications(),
            'homeowner_signatures': self._check_cedente_signature(),
            'installer': self._extract_installer(),
            'location': self._extract_location(),
            'catastral_ref': self._extract_catastral_ref(),
            'utm_coordinates': self._extract_utm_coordinates(),
            'energy_savings': self._extract_energy_savings(),
            'act_code': self._extract_act_code(),
        }
        
        return result
    
    def _extract_installer(self) -> str:
        """Extract installer name"""
        pattern = r'instalador\s+([A-ZÑ][A-ZÑ\s]+?)(?:\s+ESPA[NÑ]A)?(?:\s+\(|a\s+satisfacci[o6]n)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            installer = match.group(1).strip()
            if re.search(r'ESPA[NÑ]A', self.text[match.start():match.start()+100], re.IGNORECASE):
                installer += " ESPAÑA"
            return installer
        return "NOT FOUND"
    
    def _extract_location(self) -> str:
        """Extract location/address of installation"""
        # Pattern 1: "1. Dirección: TR PLAZA 4, 24248, URDIALES DEL PARAMO, Castilla y León."
        pattern1 = r'Direcci[oó]n:\s*(.+?\.)'
        match1 = re.search(pattern1, self.text, re.IGNORECASE | re.DOTALL)
        if match1:
            location = match1.group(1).strip()
            # Remove trailing period
            location = location.rstrip('.')
            # Clean up newlines
            location = re.sub(r'\s+', ' ', location)
            return location
        
        # Pattern 2: "en TR PLAZA 4 URDIALES DEL PARAMO Castilla y León 24248"
        pattern2 = r'en\s+([A-Z][A-Z\s\d]+?(?:PLAZA|CALLE|AVENIDA|CARRETERA)[^,\n]*?(?:Castilla y Le[o6]n\s+\d+|,\s*\d{5}))'
        match2 = re.search(pattern2, self.text, re.IGNORECASE)
        if match2:
            location = match2.group(1).strip()
            location = re.sub(r'\s+', ' ', location)
            return location
        
        # Pattern 3: "localidad de XXX Castilla y León"
        pattern3 = r'localidad de\s+(.+?)\s+(?:Castilla y Le[o6]n\s+\d+)'
        match3 = re.search(pattern3, self.text, re.IGNORECASE)
        if match3:
            location = match3.group(1).strip()
            postal_match = re.search(r'(Castilla y Le[o6]n\s+(\d+))', self.text, re.IGNORECASE)
            if postal_match:
                location += ", " + postal_match.group(1).replace('6', 'ó')
            return location
        
        return "NOT FOUND"
    
    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference"""
        # Pattern 1: "2. Referencia catastral: 2050816TM7925S0001YB"
        pattern1 = r'Referencia catastral:\s*([0-9A-Z]+)'
        match1 = re.search(pattern1, self.text, re.IGNORECASE)
        if match1:
            catastral = match1.group(1).strip()
            catastral = re.sub(r'\s+', '', catastral)
            if re.match(r'^[0-9]+[A-Z]+[0-9]+', catastral):
                return catastral
        
        # Pattern 2: "ubicación \n 5 720302TN7452S0001 KJ"
        pattern2 = r'ubicaci[o6]n\s+([0-9\s]+[A-Z]{2}\s*\d+[A-Z]\d+\s*[A-Z]{2})'
        match2 = re.search(pattern2, self.text, re.IGNORECASE)
        if match2:
            catastral = match2.group(1).strip()
            catastral = re.sub(r'\s+', '', catastral)
            return catastral
        
        return "NOT FOUND"
    
    def _extract_utm_coordinates(self) -> str:
        """Extract UTM coordinates - handles both formats"""
        # Format 1: "UTM 30, X:275624.89, Y:4741864.43"
        pattern1 = r'UTM\s+(\d+),?\s*X:\s*([\d.]+),?\s*Y:\s*([\d.\s]+)'
        match1 = re.search(pattern1, self.text, re.IGNORECASE)
        if match1:
            zone = match1.group(1)
            x = match1.group(2)
            y = match1.group(3).replace(' ', '')
            return f"X:{x} Y:{y} HUSO:{zone}"
        
        # Format 2: "• Coordenadas UTM: o X:271909.54 o Y:4694623.74 o HUSO:30"
        pattern2 = r'X:\s*([\d.]+)[^Y]*Y:\s*([\d.]+)'
        match2 = re.search(pattern2, self.text)
        if match2:
            x = match2.group(1)
            y = match2.group(2)
            # Try to find HUSO/zone
            zone_match = re.search(r'HUSO[:\s]*(\d+)', self.text, re.IGNORECASE)
            zone = zone_match.group(1) if zone_match else '30'
            return f"X:{x} Y:{y} HUSO:{zone}"
        
        return "NOT FOUND"
    
    def _extract_energy_savings(self) -> str:
        """Extract energy savings in kWh/año"""
        pattern = r'([\d.]+)\s*kWh/a[ñnrio]+(?:\s|,|\.)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return f"{match.group(1)}"
        return "NOT FOUND"
    
    def _extract_act_code(self) -> str:
        """Extract RES code (RES010, RES020)"""
        match = re.search(r'(RES0*\d{2,3})', self.text, re.IGNORECASE)
        if match:
            code = match.group(1).upper()
            code = re.sub(r'RES0+(\d)', r'RES0\1', code)
            return code
        return "NOT FOUND"
    
    def _extract_cesionario_company(self) -> str:
        """Extract Cesionario (company) name"""
        patterns = [
            r'de una parte[,\s]+(.+?)(?:con CIF|CIF|,\s*con)',
            r'Cesionario[:\s]+(.+?)(?:con CIF|CIF)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE | re.DOTALL)
            if match:
                company = match.group(1).strip()
                company = re.sub(r'\s+', ' ', company)
                if 5 < len(company) < 100:
                    return company
        
        return "NOT FOUND"
    
    def _extract_cesionario_cif(self) -> str:
        """Extract Cesionario CIF"""
        cifs = re.findall(r'\b([A-Z]\d{8})\b', self.text)
        if cifs:
            return cifs[0]
        return "NOT FOUND"
    
    def _extract_cesionario_address(self) -> str:
        """Extract Cesionario address"""
        pattern = r'domicilio[^:]*:\s*(.+?)(?:\n|,\s*CP)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return "NOT FOUND"
    
    def _extract_cesionario_representative(self) -> str:
        """Extract Cesionario representative name"""
        pattern = r'representad[oa] por\s+D[./]?\s*([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return "NOT FOUND"
    
    def _extract_cesionario_dni(self) -> str:
        """Extract Cesionario representative DNI"""
        dnis = re.findall(r'\b(\d{8}[A-Z])\b', self.text)
        if dnis:
            return dnis[0]
        return "NOT FOUND"
    
    def _check_cesionario_signature(self) -> str:
        """Check Cesionario signature"""
        if re.search(r'Firma.*Cesionario|Cesionario.*Firma', self.text, re.IGNORECASE):
            return "Present"
        return "NOT FOUND"
    
    def _extract_cedente_name(self) -> str:
        """Extract Cedente (homeowner) name"""
        # Pattern 1: "DE PAZ FRANCO QUINTILIANA, mayor de edad, con DNI"
        pattern1 = r'([A-ZÑ][A-Z\s]+?),\s*mayor de edad,\s*con DNI'
        match1 = re.search(pattern1, self.text)
        if match1:
            name = match1.group(1).strip()
            words = name.split()
            if 2 <= len(words) <= 5:
                return name
        
        # Pattern 2: Look at end of document
        lines = self.text.split('\n')
        for line in reversed(lines[-30:]):
            match = re.match(r'^([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+(?:\s+[A-ZÑ][a-zñ]+)?)\s*$', line.strip())
            if match:
                name = match.group(1)
                if not re.search(r'CARROCERA|IGLESIA|CASTILLA|URDIALES', name, re.IGNORECASE):
                    return name
        
        return "NOT FOUND"
    
    def _extract_cedente_dni(self) -> str:
        """Extract Cedente DNI"""
        dnis = re.findall(r'\b(\d{8}[A-Z])\b', self.text)
        if len(dnis) >= 2:
            return dnis[-1]
        elif len(dnis) == 1:
            return dnis[0]
        return "NOT FOUND"
    
    def _extract_cedente_address(self) -> str:
        """Extract Cedente address"""
        location = self._extract_location()
        if location != "NOT FOUND":
            return location
        
        return "NOT FOUND"
    
    def _extract_cedente_phone(self) -> str:
        """Extract Cedente phone"""
        patterns = [
            r'tel[eé]fono[:\s]*([\d\s]+)',
            r'(\+34[\s-]?\d{3}[\s-]?\d{3}[\s-]?\d{3})',
            r'(\d{9})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                phone = match.group(1).strip()
                phone = re.sub(r'\s+', '', phone)
                if len(phone) >= 9:
                    return phone
        
        return "NOT FOUND"
    
    def _extract_cedente_email(self) -> str:
        """Extract Cedente email"""
        pattern = r'[\w\.-]+@[\w\.-]+\.\w+'
        match = re.search(pattern, self.text)
        if match:
            return match.group(0)
        return "NOT FOUND"
    
    def _extract_notifications(self) -> str:
        """Extract notification preferences"""
        pattern = r'[Nn]otificaciones[:\s]+(.+?)(?:\n|$)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return "NOT FOUND"
    
    def _check_cedente_signature(self) -> str:
        """Check Cedente signatures"""
        signature_count = len(re.findall(r'Firma|Firmado|Fdo\.', self.text, re.IGNORECASE))
        if signature_count > 0:
            return f"{signature_count} signature(s) found"
        return "NOT FOUND"
