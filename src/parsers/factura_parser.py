"""
Parser for FACTURA documents
Extracts invoice details including subtotal, deductions, final amount, and installer info
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
            'homeowner_email': self._extract_email(),
            'homeowner_phone': self._extract_phone(),
            'subtotal': self._extract_subtotal(),
            'deduction': self._extract_deduction(),
            'amount': self._extract_amount(),
            # Installer info
            'installer_name': self._extract_installer_name(),
            'installer_address': self._extract_installer_address(),
            'installer_cif': self._extract_installer_cif(),
            'isolation_type': self._extract_isolation_type(),
            'isolation_thickness': self._extract_isolation_thickness(),
            's': self._extract_s(),
        }
        
        return result
    
    def _extract_invoice_number(self) -> str:
        """Extract invoice number"""
        patterns = [
            r'[Ff]actura\s*[Nn][ºo°]?\s*([A-Z0-9][\w\-]+)',
            r'[Ff]actura[:\s]+([A-Z]?\d[\w\-]+)',
            r'[Nn][ºo°]?\s*[Ff]actura[:\s]+([A-Z0-9][\w\-]+)',
            r'[Nn][ºo°]\s+([A-Z0-9][\w\-]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_invoice_date(self) -> str:
        """Extract invoice date"""
        patterns = [
            r'[Ff]echa[:\s]+([\d]{1,2}/[\d]{1,2}/[\d]{4})',
            r'\b(\d{1,2}/\d{1,2}/\d{4})\b',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_homeowner_name(self) -> str:
        """Extract customer/homeowner name"""
        patterns = [
            r'[Mm]r\.?\s+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]+)',
            r'[Ss]r[a]?\.?\s+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]+)',
            r'[Cc]liente[:\s]+([A-ZÑÁÉÍÓÚ][a-zñáéíóú]+(?:\s+[A-ZÑÁÉÍÓÚ][a-zñáéíóú]+)+)',
            r'[Nn]ombre[:\s]+([A-ZÑÁÉÍÓÚ][a-zñáéíóú]+(?:\s+[A-ZÑÁÉÍÓÚ][a-zñáéíóú]+)+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                name = match.group(1).strip()
                name = re.sub(r'\s+(CL|AV|C/|CALLE|PZ|PLAZA).*', '', name, flags=re.IGNORECASE)
                if "BONO SOCIAL" not in name.upper() and "PERCEPTORES" not in name.upper():
                    return name.strip()
        
        return "NOT FOUND"
    
    def _extract_homeowner_dni(self) -> str:
        """Extract DNI"""
        patterns = [
            r'[Dd][Nn][Ii][:\s]*(\d{7,8}[-]?[A-Z])',
            r'[Nn][Ii][Ff][:\s]*(\d{7,8}[-]?[A-Z])',
            r'\b(\d{8}[-]?[A-Z])\b',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                dni = match.group(1).replace('-', '')
                return dni
        
        return "NOT FOUND"
    
    def _extract_homeowner_address(self) -> str:
        """Extract address"""
        patterns = [
            r'((?:CL|AV|C/|CALLE|PZ|PLAZA)\s+[^\n]+\n\d{5}\s+[^\n]+)',
            r'((?:CL|AV|C/|CALLE|PZ|PLAZA)\s+[^\n]+)',
            r'[Dd]irecci[oó]n[:\s]+([^\n]+)',
            r'[Dd]omicilio[:\s]+([^\n]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                addr = match.group(1).strip()
                addr = addr.replace('\n', ', ')
                return addr
        
        return "NOT FOUND"
    
    def _extract_email(self) -> str:
        """Extract email - homeowner's email (first one found after name)"""
        # Find all emails
        emails = re.findall(r'\b([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)\b', self.text)
        
        # First email is usually the homeowner's
        if emails:
            return emails[0]
        
        return "NOT FOUND"
    
    def _extract_phone(self) -> str:
        """Extract phone number"""
        patterns = [
            r'[Tt]el(?:[eé]fono)?\.?\s*[:\s]*(\d{9})',
            r'[Tt]lf\.?\s*[:\s]*(\d{9})',
            r'[Mm][oó]vil\s*[:\s]*(\d{9})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        # Find standalone 9-digit numbers (Spanish phones)
        phones = re.findall(r'\b([69]\d{8})\b', self.text)
        if phones:
            return phones[0]
        
        return "NOT FOUND"
    
    def _extract_subtotal(self) -> str:
        """Extract subtotal amount (before deductions)"""
        patterns = [
            # "Subtotal 1319.68 €" or "Subtotal 1393.39 €"
            r'[Ss]ubtotal\s+([\d.,]+)\s*€',
            r'[Ss]ub[-]?total[:\s]*([\d.,]+)\s*€',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return f"{match.group(1)} €"
        
        return "NOT FOUND"
    
    def _extract_deduction(self) -> str:
        """Extract deduction amount (CAE subsidies) - negative number"""
        patterns = [
            # Negative number in brackets like "[1392.39€" or "-1318.68€"
            r'\[?([-]?[\d.,]+)\s*€?\s*\+?',
            r'(-[\d.,]+)\s*€',
        ]
        
        # Look for negative amounts
        matches = re.findall(r'(-[\d.,]+)\s*€', self.text)
        if matches:
            # Return the largest negative deduction
            return f"{matches[0]} €"
        
        # Look for bracketed amounts like [1392.39€
        match = re.search(r'\[([\d.,]+)\s*€', self.text)
        if match:
            return f"-{match.group(1)} €"
        
        return "NOT FOUND"
    
    def _extract_amount(self) -> str:
        """Extract total amount (A pagar)"""
        patterns = [
            r'[Aa]\s*[Pp]agar[:\s]*([\d.,]+)\s*€',
            r'(?<![Ss]ub)[Tt]otal[:\s]+([\d.,]+)\s*€',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return f"{match.group(1)} €"
        
        return "NOT FOUND"
    
    def _extract_installer_name(self) -> str:
        """Extract installer company name from bottom of invoice"""
        patterns = [
            # "ECORENOVA ESPANA;" or "ETL Aislamiento;"
            r'(ECORENOVA\s*ESPA[NÑ]A)',
            r'(ETL\s*[Aa]islamiento)',
            r'^([A-Z][A-Z\s]+);',  # Company name ending with semicolon
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.MULTILINE)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_installer_address(self) -> str:
        """Extract installer address from bottom of invoice"""
        patterns = [
            # "C TUSET, NUM 20 PLANTA 8, PUERTA, 8 08006 BARCELONA"
            r'([Cc]\s+[A-Z][A-Z\s,]+\d{5}\s+[A-Z]+)',
            r'(CALLE\s+[A-Z][A-Z\s,]+\d{5}\s+[A-Z]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1).strip()
        
        return "NOT FOUND"
    
    def _extract_installer_cif(self) -> str:
        """Extract installer CIF"""
        patterns = [
            r'C\.?I\.?F\.?[:\s]*([AB]\d{8})',
            r'CIF[:\s]*([AB]\d{8})',
            r'\b([AB]\d{8})\b',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_s(self) -> str:
        """Extract surface area (S) in m²"""
        # First, try specific patterns
        patterns = [
            r'[Ss]uperficie[:\s]*([\d.,]+)\s*m[²2]?',
            r'[Ss]uperficie\s+total[:\s]*([\d.,]+)\s*m[²2]?',
            r'[Mm]etros\s+cuadrados[:\s]*([\d.,]+)',
            r'[Ss]\s*[:=]\s*([\d.,]+)',
            r'[Aa]rea[:\s]*([\d.,]+)\s*m[²2]?',
            r'(\d{2,4})\s*m[²2]',
            r'(\d{2,4})\s*metros\s+cuadrados',
            r'aislamiento\s+(\d{2,3}(?:\.\d)?)',  # 50.0 after aislamiento
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                s_value = match.group(1).replace(',', '.')
                try:
                    num = float(s_value)
                    if 10 <= num <= 1000:  # reasonable surface range
                        return str(int(num))  # return as integer
                except ValueError:
                    continue
        
        # Fallback: look for any number near m²
        match = re.search(r'(\d{2,4})\s*(?:m[²2]|metros?\s+cuadrados?)', self.text, re.IGNORECASE)
        if match:
            s_value = match.group(1)
            try:
                num = int(s_value)
                if 10 <= num <= 1000:
                    return str(num)
            except ValueError:
                pass
        
        return "NOT FOUND"
    
    def _extract_isolation_type(self) -> str:
        """Extract isolation type: return only SOPLADO or ROLLO (or NOT FOUND)."""
        t = self.text or ""
        # Prefer explicit "tipo Soplado/Rollo"
        m = re.search(r'tipo\s+(Soplado|Rollo)', t, re.IGNORECASE)
        if m:
            return m.group(1).upper()

        # Sometimes the word appears standalone
        m = re.search(r'\b(Soplado|Rollo)\b', t, re.IGNORECASE)
        if m:
            return m.group(1).upper()

        # Fallback: common brand URSA implies soplado in our dataset
        if 'URSA' in t.upper():
            return 'SOPLADO'

        return "NOT FOUND"
    
    def _extract_isolation_thickness(self) -> str:
        """Extract isolation thickness in mm"""
        patterns = [
            r'[Ee]spesor\s+aislante[:\s]*([\d.,]+)\s*mm',
            r'[Gg]rosor\s+aislante[:\s]*([\d.,]+)\s*mm',
            r'aislante\s+de\s+([\d.,]+)\s*mm',
            r'(\d{2,3})\s*mm',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                thickness = match.group(1).replace(',', '.')
                try:
                    num = float(thickness)
                    if 10 <= num <= 500:  # reasonable thickness range
                        return f"{int(num)} mm"
                except ValueError:
                    continue
        
        return "NOT FOUND"