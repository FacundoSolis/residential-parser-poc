"""
Base parser class for all document types
"""

import pdfplumber
from typing import Dict, Any, Optional
from pathlib import Path


class BaseDocumentParser:
    """Base class for all residential document parsers"""
    
    def __init__(self, pdf_path: str):
        self.pdf_path = Path(pdf_path)
        self.text = ""
        self.document_type = self._identify_document_type()
        
    def _identify_document_type(self) -> str:
        """Identify document type from filename"""
        filename = self.pdf_path.name.upper()
        
        if 'CONTRATO' in filename:
            return 'CONTRATO_CESION_AHORROS'
        elif 'FICHA' in filename:
            return 'FICHA_RES020'
        elif 'DECLARACION' in filename:
            return 'DECLARACION_RESPONSABLE'
        elif 'FACTURA' in filename:
            return 'FACTURA'
        elif 'FOTOGRAFICO' in filename:
            return 'INFORME_FOTOGRAFICO'
        elif 'CERTIFICADO' in filename:
            return 'CERTIFICADO_INSTALADOR'
        elif 'CEE FINAL' in filename:
            return 'CEE_FINAL'
        elif 'REGISTRO' in filename:
            return 'REGISTRO_CEE'
        elif 'DNI' in filename:
            return 'DNI'
        else:
            return 'UNKNOWN'
    
    def extract_text(self) -> str:
        """Extract text from PDF"""
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                pages = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        pages.append(text)
                self.text = '\n'.join(pages)
            return self.text
        except Exception as e:
            print(f"Error extracting text from {self.pdf_path}: {e}")
            return ""
    
    def parse(self) -> Dict[str, Any]:
        """
        Parse document and extract fields
        Override in subclasses
        """
        raise NotImplementedError("Subclasses must implement parse()")
    
    def _parse_date(self, date_str: str) -> Optional[str]:
        """Parse date in DD/MM/YYYY format"""
        # TODO: Implement date parsing
        return date_str