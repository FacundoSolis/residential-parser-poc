"""
Base parser with automatic OCR support for scanned PDFs and images
"""
import re
from pathlib import Path
from typing import Optional

import pdfplumber
import pytesseract
pytesseract.pytesseract.tesseract_cmd = "/opt/local/bin/tesseract"



class BaseDocumentParser:
    """Base class for document parsers with automatic OCR fallback"""

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.text = ""
        self._ocr_available: Optional[bool] = None

    def _check_ocr_available(self) -> bool:
        """Check if OCR dependencies are available (tesseract + poppler)."""
        if self._ocr_available is not None:
            return self._ocr_available

        try:
            import pytesseract  # noqa: F401
            from pdf2image import convert_from_path  # noqa: F401

            version = pytesseract.get_tesseract_version()
            print("✅ Tesseract version:", version)

            self._ocr_available = True
        except Exception as e:
            print("❌ OCR not available:", e)
            self._ocr_available = False

        return self._ocr_available

    def _extract_with_ocr_pdf(self) -> str:
        """Extract text using OCR for scanned PDFs."""
        try:
            from pdf2image import convert_from_path

            # Heurística: si es DNI (por nombre de archivo), subimos DPI y ajustamos OCR
            is_dni = "dni" in self.file_path.name.lower()

            dpi = 450 if is_dni else 300
            images = convert_from_path(
                str(self.file_path),
                dpi=dpi,
                poppler_path="/opt/local/bin",
            )

            text_parts = []
            for i, image in enumerate(images):
                if is_dni:
                    # PSM 6 suele ir bien para bloques de texto
                    config = "--oem 1 --psm 6"
                    page_text = pytesseract.image_to_string(image, lang="spa", config=config)
                else:
                    page_text = pytesseract.image_to_string(image, lang="spa")

                text_parts.append(page_text)
                print(f"  OCR page {i+1}: {len(page_text)} chars")

            return "\n".join(text_parts)

        except Exception as e:
            print(f"  OCR PDF failed: {e}")
            return ""


    def _extract_with_ocr_image(self) -> str:
        """Extract text using OCR for images (.jpg/.jpeg/.png)."""
        try:
            import pytesseract
            from PIL import Image

            img = Image.open(self.file_path)
            return pytesseract.image_to_string(img, lang="spa")

        except Exception as e:
            print(f"  OCR image failed: {e}")
            return ""

    def extract_text(self) -> str:
        """Extract text from PDF/image, with automatic OCR fallback."""
        if self.text:
            return self.text

        ext = self.file_path.suffix.lower()

        # --- Images: OCR only ---
        if ext in {".jpg", ".jpeg", ".png"}:
            if self._check_ocr_available():
                self.text = self._extract_with_ocr_image()
            else:
                print("  ⚠️ Image requires OCR but OCR not available.")
                self.text = ""
            return self.text

        # --- PDFs: try text layer first (pdfplumber) ---
        if ext == ".pdf":
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

            # --- If scanned/empty: OCR fallback ---
            if len(self.text.strip()) < 50:
                if self._check_ocr_available():
                    print("  PDF appears scanned, using OCR...")
                    self.text = self._extract_with_ocr_pdf()
                else:
                    print("  ⚠️ PDF is scanned but OCR not available.")
                    self.text = ""

            if self.text:
                print(f"✓ Extracted {len(self.text)} chars from {self.file_path.name}")
            else:
                print(f"✗ No text extracted from {self.file_path.name}")

            return self.text

        # Unsupported type
        print(f"  ⚠️ Unsupported file type: {ext}")
        self.text = ""
        return self.text

    def parse(self) -> dict:
        """Override in subclasses"""
        raise NotImplementedError("Subclasses must implement parse()")

    def _clean_text(self, text: str) -> str:
        if not text:
            return ""
        # colapsa espacios
        text = re.sub(r"[ \t]+", " ", text)
        # quita líneas de 1-2 caracteres (basura OCR)
        lines = []
        for ln in text.splitlines():
            s = ln.strip()
            if len(s) <= 2 and re.fullmatch(r"[A-Za-zÁÉÍÓÚÑ]", s or ""):
                continue
            lines.append(ln)
        text = "\n".join(lines)
        # reduce saltos excesivos
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

