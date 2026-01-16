"""
Parser for CERTIFICADO INSTALADOR documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class CertificadoParser(BaseDocumentParser):
    """Parser for Certificado Instalador"""
    
    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Certificado"""
        self.extract_text()
        
        result = {
            'document_type': 'CERTIFICADO_INSTALADOR',
            'act_code': self._extract_act_code(),
            'energy_savings': self._extract_energy_savings(),
            'start_date': self._extract_start_date(),
            'finish_date': self._extract_finish_date(),
            'address': self._extract_property_address(),
            'catastral_ref': self._extract_catastral_ref(),
            'lifespan': self._extract_lifespan(),
            'surface': self._extract_surface(),
            'climatic_zone': self._extract_climatic_zone(),
        }
        
        return result
    
    def _extract_act_code(self) -> str:
        """Extract RES code"""
        match = re.search(r'(RES0*\d{2,3})', self.text, re.IGNORECASE)
        if match:
            code = match.group(1).upper()
            code = re.sub(r'RES0+(\d)', r'RES0\1', code)
            return code
        return "NOT FOUND"
    
    def _extract_energy_savings(self) -> str:
        """Extract energy savings"""
        patterns = [
            r'Ahorro.*?AE\s+([\d.]+)',
            r'AE\s+([\d.]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return f"{match.group(1)}"
        return "NOT FOUND"
    
    def _extract_start_date(self) -> str:
        """Extract start date"""
        patterns = [
            r'inici[oó]\s+el\s+(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})',
            r'[Ff]echa.*?inicio[:\s]+([\d/]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                if len(match.groups()) == 3:
                    day, month, year = match.groups()
                    months = {
                        'enero': '01', 'febrero': '02', 'marzo': '03',
                        'abril': '04', 'mayo': '05', 'junio': '06',
                        'julio': '07', 'agosto': '08', 'septiembre': '09',
                        'octubre': '10', 'noviembre': '11', 'diciembre': '12'
                    }
                    month_num = months.get(month.lower(), '00')
                    return f"{day}/{month_num}/{year}"
                else:
                    return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_finish_date(self) -> str:
        """Extract finish date - handle newlines with \s+"""
        patterns = [
            r'finaliz[oó].*?el\s+(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})',
            r'[Ff]echa.*?fin[:\s]+([\d/]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE | re.DOTALL)
            if match:
                if len(match.groups()) == 3:
                    day, month, year = match.groups()
                    months = {
                        'enero': '01', 'febrero': '02', 'marzo': '03',
                        'abril': '04', 'mayo': '05', 'junio': '06',
                        'julio': '07', 'agosto': '08', 'septiembre': '09',
                        'octubre': '10', 'noviembre': '11', 'diciembre': '12'
                    }
                    month_num = months.get(month.lower(), '00')
                    return f"{day}/{month_num}/{year}"
                else:
                    return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_property_address(self) -> str:
        """Extract property address"""
        pattern = r'situado en\s+(.+?)(?:\.|La referencia|$)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            address = match.group(1).strip()
            address = re.sub(r'\s+', ' ', address)
            return address
        return "NOT FOUND"
    
    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference"""
        pattern = r'referencia catastral[^.]*es\s+([0-9A-Z]+)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            catastral = match.group(1).strip()
            catastral = re.sub(r'\s+', '', catastral)
            if re.match(r'^[0-9]+[A-Z]+[0-9]+[A-Z]+[0-9]+[A-Z]+$', catastral):
                return catastral
        return "NOT FOUND"
    
    def _extract_lifespan(self) -> str:
        """Extract lifespan in years"""
        pattern = r'duración.*?actuación.*?(\d+)\s*años'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return f"{match.group(1)}"
        return "NOT FOUND"
    
    def _extract_surface(self) -> str:
        """Extract surface in m2"""
        pattern = r'superficie tratada.*?([\d.]+)\s*m'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return f"{match.group(1)}"
        return "NOT FOUND"
    
    def _extract_climatic_zone(self) -> str:
        """Extract climatic zone"""
        pattern = r'zona climática.*?es.*?([A-E]\d)'
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).upper()
        return "NOT FOUND"
