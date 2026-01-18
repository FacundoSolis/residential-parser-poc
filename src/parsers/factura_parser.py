"""
Parser for FACTURA documents
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
            'amount': self._extract_amount(),
        }
        
        return result
    
    def _extract_invoice_number(self) -> str:
        """Extract invoice number"""
        patterns = [
            # "Factura Nº 15-10-2025-ND005117"
            r'[Ff]actura\s*[Nn][ºo°]?\s*([A-Z0-9][\w\-]+)',
            # "Factura: F202502927"
            r'[Ff]actura[:\s]+([A-Z]?\d[\w\-]+)',
            # "Nº Factura: XXX"
            r'[Nn][ºo°]?\s*[Ff]actura[:\s]+([A-Z0-9][\w\-]+)',
            # "N° XXX" or "Nº XXX"
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
            # "Fecha: 15/10/2025"
            r'[Ff]echa[:\s]+([\d]{1,2}/[\d]{1,2}/[\d]{4})',
            # Any date format dd/mm/yyyy
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
            # "Mr FERRERO PEREZ EVENCIO" or "Mr. FERRERO PEREZ EVENCIO"
            r'[Mm]r\.?\s+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]+)',
            # "Sra. NAME" or "Sr. NAME"
            r'[Ss]r[a]?\.?\s+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]+)',
            # "Cliente: Name"
            r'[Cc]liente[:\s]+([A-ZÑÁÉÍÓÚ][a-zñáéíóú]+(?:\s+[A-ZÑÁÉÍÓÚ][a-zñáéíóú]+)+)',
            # "Nombre: Name"  
            r'[Nn]ombre[:\s]+([A-ZÑÁÉÍÓÚ][a-zñáéíóú]+(?:\s+[A-ZÑÁÉÍÓÚ][a-zñáéíóú]+)+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                name = match.group(1).strip()
                # Clean up: remove trailing address parts
                name = re.sub(r'\s+(CL|AV|C/|CALLE|PZ|PLAZA).*', '', name, flags=re.IGNORECASE)
                return name.strip()
        
        return "NOT FOUND"
    
    def _extract_homeowner_dni(self) -> str:
        """Extract DNI"""
        patterns = [
            # "DNI:10153878E" or "DNI: 10153878E"
            r'[Dd][Nn][Ii][:\s]*(\d{7,8}[A-Z])',
            # "NIF: 10153878E"
            r'[Nn][Ii][Ff][:\s]*(\d{7,8}[A-Z])',
            # Standalone DNI pattern
            r'\b(\d{8}[A-Z])\b',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_homeowner_address(self) -> str:
        """Extract address - lines after name, before Tel/Email"""
        # Try to find address between name and contact info
        patterns = [
            # Address line starting with CL, AV, C/, CALLE, PZ, PLAZA followed by postal code line
            r'((?:CL|AV|C/|CALLE|PZ|PLAZA)\s+[^\n]+\n\d{5}\s+[^\n]+)',
            # "Dirección: XXX"
            r'[Dd]irecci[oó]n[:\s]+([^\n]+)',
            # "Domicilio: XXX"
            r'[Dd]omicilio[:\s]+([^\n]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                addr = match.group(1).strip()
                # Clean up newlines
                addr = addr.replace('\n', ', ')
                return addr
        
        return "NOT FOUND"
    
    def _extract_email(self) -> str:
        """Extract email"""
        patterns = [
            r'[Ee]mail\s*[:\s]*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)',
            r'[Cc]orreo\s*[:\s]*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)',
            r'\b([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)\b',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_phone(self) -> str:
        """Extract phone number"""
        patterns = [
            r'[Tt]el(?:[eé]fono)?\.?\s*[:\s]*(\d{9})',
            r'[Tt]lf\.?\s*[:\s]*(\d{9})',
            r'[Mm][oó]vil\s*[:\s]*(\d{9})',
            r'\b(6\d{8})\b',  # Spanish mobile starting with 6
            r'\b(9\d{8})\b',  # Spanish landline starting with 9
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return match.group(1)
        
        return "NOT FOUND"
    
    def _extract_amount(self) -> str:
        """Extract total amount"""
        patterns = [
            # "A pagar 1.21€" or "A pagar: 1.21 €"
            r'[Aa]\s*[Pp]agar[:\s]*([\d.,]+)\s*€',
            # "Total 1 €"
            r'[Tt]otal[:\s]+([\d.,]+)\s*€',
            # "Importe: XXX €"
            r'[Ii]mporte[:\s]+([\d.,]+)\s*€',
            # "TOTAL: XXX €"
            r'TOTAL[:\s]+([\d.,]+)\s*€',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, self.text)
            if match:
                return f"{match.group(1)} €"
        
        return "NOT FOUND"