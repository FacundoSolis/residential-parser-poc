"""
Base parser with automatic OCR support for scanned PDFs and images
"""
import re
from pathlib import Path
from typing import Optional
import os
import shutil

import pdfplumber
import pytesseract


def _configure_binaries():
    # Tesseract
    tess = os.getenv("TESSERACT_CMD") or shutil.which("tesseract") or "/usr/bin/tesseract"
    pytesseract.pytesseract.tesseract_cmd = tess

    # Poppler (optional). If pdftoppm exists in PATH, pdf2image can work without poppler_path.
    poppler = os.getenv("POPPLER_PATH")
    return poppler if poppler else None


POPPLER_PATH = _configure_binaries()


class BaseDocumentParser:
    """Base class for document parsers with automatic OCR fallback"""

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.text = ""
        self._ocr_available: Optional[bool] = None

    def _check_ocr_available(self) -> bool:
        if self._ocr_available is not None:
            return self._ocr_available

        try:
            from pdf2image import convert_from_path  # noqa: F401

            tess = shutil.which("tesseract")
            pdftoppm = shutil.which("pdftoppm")

            if not tess:
                raise RuntimeError("tesseract not found in PATH")
            if not pdftoppm and not POPPLER_PATH:
                raise RuntimeError("poppler (pdftoppm) not found in PATH and POPPLER_PATH not set")

            _ = pytesseract.get_tesseract_version()
            self._ocr_available = True
        except Exception as e:
            print("❌ OCR not available:", e)
            self._ocr_available = False

        return self._ocr_available

    def _extract_with_ocr_pdf(self) -> str:
        """Extract text using OCR for scanned PDFs."""
        try:
            from pdf2image import convert_from_path

            is_dni = "dni" in self.file_path.name.lower()
            dpi = 450 if is_dni else 300

            images = convert_from_path(
                str(self.file_path),
                dpi=dpi,
                poppler_path=POPPLER_PATH,  # can be None
            )

            text_parts = []
            for i, image in enumerate(images):
                if is_dni:
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
                self.text = self._clean_text(self.text)  # ✅ clean
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
                    self.text = self._clean_text(self.text)  # ✅ clean
            except Exception as e:
                print(f"  pdfplumber error: {e}")
                self.text = ""

            # --- If scanned/empty: OCR fallback ---
            if len(self.text.strip()) < 50 or (("contrato" in self.file_path.name.lower() or "declaracion" in self.file_path.name.lower()) and len(self.text.strip()) < 2000):
                if self._check_ocr_available():
                    print("  PDF appears scanned, using OCR...")
                    self.text = self._extract_with_ocr_pdf()
                    self.text = self._clean_text(self.text)  # ✅ clean
                else:
                    print("  ⚠️ PDF is scanned but OCR not available.")
                    self.text = ""

            if self.text:
                print(f"✓ Extracted {len(self.text)} chars from {self.file_path.name}")
            else:
                print(f"✗ No text extracted from {self.file_path.name}")

            return self.text

        print(f"  ⚠️ Unsupported file type: {ext}")
        self.text = ""
        return self.text

    def parse(self) -> dict:
        raise NotImplementedError("Subclasses must implement parse()")

    def _clean_text(self, text: str) -> str:
        if not text:
            return ""
        text = re.sub(r"[ \t]+", " ", text)
        lines = []
        for ln in text.splitlines():
            s = ln.strip()
            if len(s) <= 2 and re.fullmatch(r"[A-Za-zÁÉÍÓÚÑ]", s or ""):
                continue
            lines.append(ln)
        text = "\n".join(lines)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    def _detect_isolation_type(self, text: str):
        """Detect isolation type (SOPLADO or ROLLO) and whether it was inferred by brand.

        Returns a tuple `(type_str, inferred_by_brand)` where `type_str` is the
        normalized type or "NOT FOUND" and `inferred_by_brand` is a bool.
        """
        t = (text or "").replace('-', ' ')
        # 1) explicit "tipo Soplado/Rollo"
        m = re.search(r"tipo\s+(soplado|rollo)", t, re.IGNORECASE)
        if m:
            return m.group(1).upper(), False

        # 2) standalone words
        m = re.search(r"\b(soplado|rollo)\b", t, re.IGNORECASE)
        if m:
            return m.group(1).upper(), False

        # 3) brand/product heuristics (fallbacks) - map known brands to a likely type
        up = t.upper()
        if 'URSA' in up:
            return 'SOPLADO', True
        if 'ISOVER' in up or 'ISOLENE' in up or 'SAINT GOBAIN'.upper() in up:
            return 'SOPLADO', True

        return 'NOT FOUND', False
