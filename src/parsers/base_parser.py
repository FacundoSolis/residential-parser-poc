"""
Base parser with automatic OCR support for scanned PDFs
"""
import pdfplumber
from pathlib import Path
from typing import Optional
import re


class BaseDocumentParser:
    """Base class for document parsers with automatic OCR fallback"""
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.text = ""
        self._ocr_available = None
    
    def _check_ocr_available(self) -> bool:
        """Check if OCR dependencies are available"""
        if self._ocr_available is not None:
            return self._ocr_available
        
        try:
            import pytesseract
            from pdf2image import convert_from_path
            # Test tesseract is installed
            pytesseract.get_tesseract_version()
            self._ocr_available = True
        except Exception:
            self._ocr_available = False
        
        return self._ocr_available
    
    def _extract_with_ocr(self) -> str:
        """Extract text using OCR (for scanned PDFs)"""
        try:
            import pytesseract
            from pdf2image import convert_from_path
            
            # Convert PDF pages to images
            images = convert_from_path(str(self.file_path), dpi=300)
            
            text_parts = []
            for i, image in enumerate(images):
                # Use Spanish language for better recognition
                page_text = pytesseract.image_to_string(image, lang='spa')
                text_parts.append(page_text)
                print(f"  OCR page {i+1}: {len(page_text)} chars")
            
            return "\n".join(text_parts)
        
        except Exception as e:
            print(f"  OCR failed: {e}")
            return ""
    
    def extract_text(self) -> str:
        """Extract text from PDF, with automatic OCR fallback for scanned documents"""
        if self.text:
            return self.text
        
        # First try pdfplumber (for digital PDFs)
        try:
            with pdfplumber.open(self.file_path) as pdf:
                text_parts = []
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    text_parts.append(page_text)
                
                self.text = "\n".join(text_parts)
        except Exception as e:
            print(f"  pdfplumber error: {e}")
            self.text = ""
        
        # Check if we got meaningful text
        # If text is too short or empty, try OCR
        if len(self.text.strip()) < 100:
            if self._check_ocr_available():
                print(f"  PDF appears scanned, using OCR...")
                self.text = self._extract_with_ocr()
            else:
                print(f"  ⚠️  PDF is scanned but OCR not available. Install: brew install tesseract tesseract-lang poppler && pip install pytesseract pdf2image")
        
        if self.text:
            print(f"✓ Extracted {len(self.text)} chars from {self.file_path.name}")
        else:
            print(f"✗ No text extracted from {self.file_path.name}")
        
        return self.text
    
    def parse(self) -> dict:
        """Override in subclasses"""
        raise NotImplementedError("Subclasses must implement parse()")
    
    def _clean_text(self, text: str) -> str:
        """Clean extracted text"""
        if not text:
            return ""
        # Remove multiple spaces
        text = re.sub(r' +', ' ', text)
        # Remove multiple newlines
        text = re.sub(r'\n{3,}', '\n\n', text)
        return text.strip()