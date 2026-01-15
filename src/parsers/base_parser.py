"""
Base parser class for all document types
"""

import pdfplumber
import re
from typing import Dict, Any, Optional
from pathlib import Path
from datetime import datetime


class BaseDocumentParser:
    """Base class for all residential document parsers"""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = Path(pdf_path)
        self.text = ""
        
    def extract_text(self) -> str:
        """Extract text from all pages of PDF"""
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                pages = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        pages.append(text)
                self.text = '\n'.join(pages)
            
            print(f"✓ Extracted {len(self.text)} chars from {self.pdf_path.name}")
            return self.text
            
        except Exception as e:
            print(f"✗ Error extracting from {self.pdf_path.name}: {e}")
            return ""
    
    def parse(self) -> Dict[str, Any]:
        """
        Parse document and extract fields
        Override in subclasses
        """
        self.extract_text()
        return {}
    
    def _parse_date(self, date_str: str) -> Optional[str]:
        """Parse date in DD/MM/YYYY format"""
        if not date_str:
            return None
        
        # Try DD/MM/YYYY format
        match = re.search(r'(\d{2})/(\d{2})/(\d{4})', date_str)
        if match:
            return f"{match.group(1)}/{match.group(2)}/{match.group(3)}"
        
        return None
    
    def _extract_field_after_label(self, label: str) -> Optional[str]:
        """Extract value that appears after a label"""
        pattern = f"{label}[:\\s]+(.+?)(?=\\n|$)"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return None
    
    def _extract_cif(self) -> Optional[str]:
        """Extract CIF (Spanish company tax ID)"""
        # CIF format: Letter + 8 digits
        match = re.search(r'\b[A-Z]\d{8}\b', self.text)
        if match:
            return match.group(0)
        return None
    
    def _extract_dni(self) -> Optional[str]:
        """Extract DNI (Spanish ID number)"""
        # DNI format: 8 digits + Letter
        match = re.search(r'\b\d{8}[A-Z]\b', self.text)
        if match:
            return match.group(0)
        return None
