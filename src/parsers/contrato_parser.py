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
        pattern = r'localidad de\s+(.+?)\s+(?:Castilla y Le[o6]n\s+\d+)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            location = match.group(1).strip()
            postal_match = re.search(r'(Castilla y Le[o6]n\s+(\d+))', self.text, re.IGNORECASE)
            if postal_match:
                location += ", " + postal_match.group(1).replace('6', 'ó')
            return location
        return "NOT FOUND"
    
    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference - number appears AFTER 'ubicación' on next line"""
        # Pattern: "referencia catastral de su ubicación \n 5 720302TN7452S0001 KJ"
        # The catastral ref is on the line AFTER "ubicación"
        
        # Find "ubicación" then capture next line with catastral format
        pattern = r'ubicaci[o6]n\s+([0-9\s]+[A-Z]{2}\s*\d+[A-Z]\d+\s*[A-Z]{2})'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            catastral = match.group(1).strip()
            # Remove ALL spaces
            catastral = re.sub(r'\s+', '', catastral)
            return catastral
        
        return "NOT FOUND"
    
    def _extract_utm_coordinates(self) -> str:
        """Extract UTM coordinates"""
        pattern = r'UTM\s+(\d+),?\s*X:\s*([\d.]+),?\s*Y:\s*([\d.\s]+)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            zone = match.group(1)
            x = match.group(2)
            y = match.group(3).replace(' ', '')
            return f"UTM {zone}, X:{x}, Y:{y}"
        return "NOT FOUND"
    
    def _extract_energy_savings(self) -> str:
        """Extract energy savings in kWh/año"""
        pattern = r'([\d.]+)\s*kWh/a[ñnrio]+(?:\s|,|\.)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return f"{match.group(1)} kWh/año"
        return "NOT FOUND"
    
    def _extract_cesionario_company(self) -> str:
        """Extract Cesionario (company) name - likely in page 1"""
        patterns = [
            r'de una parte[,\s]+(.+?)(?:con CIF|CIF|,\s*con)',
            r'Cesionario[:\s]+(.+?)(?:con CIF|CIF)',
            r'representada por[:\s]+D[./]?\s*([A-ZÑ\s]+)(?:con DNI)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE | re.DOTALL)
            if match:
                company = match.group(1).strip()
                company = re.sub(r'\s+', ' ', company)
                if 5 < len(company) < 100:
                    return company
        
        return "NOT FOUND - Check Page 1"
    
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
        lines = self.text.split('\n')
        for line in reversed(lines[-30:]):
            match = re.match(r'^([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+(?:\s+[A-ZÑ][a-zñ]+)?)\s*$', line.strip())
            if match:
                name = match.group(1)
                if not re.search(r'CARROCERA|IGLESIA|CASTILLA', name, re.IGNORECASE):
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
        """Extract Cedente address (same as location)"""
        return self._extract_location()
    
    def _extract_cedente_phone(self) -> str:
        """Extract Cedente phone"""
        patterns = [
            r'[Tt]el[eé]fono[:\s]*([\d\s-]+)',
            r'(\+34[\s-]?\d{3}[\s-]?\d{3}[\s-]?\d{3})',
            r'(\d{3}[\s-]\d{3}[\s-]\d{3})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
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
