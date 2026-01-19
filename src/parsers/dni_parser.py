"""
Parser for DNI documents

✅ Robust against OCR noise:
- misread digits (O->0, I/L->1)
- multiple layouts (APELLIDOS/NOMBRE, APELLIDOS Y NOMBRE, 1 APELLIDO/2 APELLIDO/NOMBRE)
- fallback line heuristics
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class DniParser(BaseDocumentParser):
    """Parser for DNI"""

    def parse(self) -> Dict[str, Any]:
        self.extract_text()
        return {
            "document_type": "DNI",
            "dni_number": self._extract_dni_number(),
            "name": self._extract_name(),
        }

    def _normalize_digits(self, t: str) -> str:
        t = t.upper()
        t = t.replace(" ", "").replace("-", "").replace(".", "").replace(":", "")
        t = t.replace("O", "0")
        t = t.replace("I", "1")
        t = t.replace("L", "1")
        return t

    def _extract_dni_number(self) -> str:
        """Extract DNI/NIF number (8 digits + letter)."""
        if not self.text:
            return "NOT FOUND"
        t = self._normalize_digits(self.text)

        m = re.search(r"\b(\d{8}[A-Z])\b", t)
        if m:
            return m.group(1)

        m = re.search(r"(\d{8}[A-Z])", t)
        if m:
            return m.group(1)

        return "NOT FOUND"

    def _extract_name(self) -> str:
        if not self.text:
            return "NOT FOUND"

        t = self.text.strip()

        patterns = [
            r"APELLIDOS\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{4,})\s+NOMBRE\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{2,})",
            r"NOMBRE\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{2,})\s+APELLIDOS\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{4,})",
            r"APELLIDOS\s+Y\s+NOMBRE\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{8,})",
            r"1\s*APELLIDO\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{2,})\s+2\s*APELLIDO\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{2,})\s+NOMBRE\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ\s]{2,})",
        ]

        for p in patterns:
            m = re.search(p, t, re.IGNORECASE)
            if not m:
                continue

            if len(m.groups()) == 1:
                name = self._clean_name(m.group(1))
            elif len(m.groups()) == 2:
                a = self._clean_name(m.group(1))
                b = self._clean_name(m.group(2))
                name = f"{b} {a}".strip() if p.startswith("NOMBRE") else f"{a} {b}".strip()
            else:
                a1 = self._clean_name(m.group(1))
                a2 = self._clean_name(m.group(2))
                nom = self._clean_name(m.group(3))
                name = f"{a1} {a2} {nom}".strip()

            # FILTRO FINAL
            return name if not self._is_garbage_name(name) else "NOT FOUND"

        # Line heuristics
        lines = [self._clean_name(ln) for ln in t.splitlines() if ln.strip()]

        caps_lines = [ln for ln in lines if re.fullmatch(r"[A-ZÑÁÉÍÓÚ\s]{10,}", ln)]
        if caps_lines:
            cand = max(caps_lines, key=len)
            return cand if not self._is_garbage_name(cand) else "NOT FOUND"

        return "NOT FOUND"


    def _clean_name(self, name: str) -> str:
        name = re.sub(r"\s+", " ", (name or "")).strip()
        return name

    def _is_garbage_name(self, name: str) -> bool:
        if not name or name == "NOT FOUND":
            return True
        up = name.upper()

        # si contiene etiquetas típicas => basura OCR
        if re.search(r"\b(APELLIDOS|NOMBRE|DOCUMENTO|NIF|DNI|SEXO|NACIONALIDAD|FECHA|CADUCIDAD)\b", up):
            return True

        # muy corto o muy pocas letras
        if len(up) < 10 or len(up.split()) < 2:
            return True

        # demasiadas letras sueltas tipo "UN VI Y E"
        tokens = up.split()
        short_tokens = sum(1 for tok in tokens if len(tok) <= 2)
        if short_tokens >= 2:
            return True

        return False

