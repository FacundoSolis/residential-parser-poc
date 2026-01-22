"""
Parser for DNI documents

✅ Robust against OCR noise:
- misread digits (O->0)
- I/L confusion ONLY inside the 8-digit numeric part (never touch final letter)
- last letter misread (G->6)
- MRZ / IDESP fallback (embedded DNI like IDESP...13103004L)
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

    # ------------------------
    # Helpers
    # ------------------------

    def _normalize_common(self, t: str) -> str:
        """Upper + remove separators. IMPORTANT: do NOT replace L->1 here."""
        t = (t or "").upper()
        t = t.replace(" ", "").replace("-", "").replace(".", "").replace(":", "")
        return t

    def _normalize_numeric_part(self, eight: str) -> str:
        """Normalize only the numeric part (8 chars) allowing OCR swaps."""
        eight = (eight or "").upper()
        eight = eight.replace("O", "0").replace("I", "1").replace("L", "1")
        return eight

    def _is_valid_spanish_dni(self, dni: str) -> bool:
        dni = (dni or "").strip().upper().replace(" ", "")
        if not re.match(r"^\d{8}[A-Z]$", dni):
            return False
        letters = "TRWAGMYFPDXBNJZSQVHLCKE"
        num = int(dni[:8])
        return dni[-1] == letters[num % 23]

    # ------------------------
    # DNI extractor
    # ------------------------

    def _extract_dni_number(self) -> str:
        """
        Goals:
        - Catch clean DNI anywhere: 13103004L (even inside IDESP...13103004L)
        - OCR tolerant: O/I/L may appear inside the 8 digits only
        - Last letter: accept '6' as 'G' (common OCR)
        - Never convert final 'L' -> '1' (that was the bug)
        """
        if not self.text:
            return "NOT FOUND"

        raw = (self.text or "").upper()

        # 1) Fast path: find any clean DNI (works for embedded IDESP...13103004L)
        # NOTE: this does not remove chars, so it can match inside long tokens.
        m = re.search(r"(\d{8}[A-Z])", raw)
        if m and self._is_valid_spanish_dni(m.group(1)):
            return m.group(1)

        # 2) OCR-tolerant scan:
        # - Numeric part may contain O/I/L
        # - Last char may be A-Z or 6 (G) or I/L/1 (noise)
        t = self._normalize_common(raw)

        for m in re.finditer(r"([0-9OIL]{8})([A-Z6IL1])", t):
            raw_num, raw_last = m.groups()

            num = self._normalize_numeric_part(raw_num)

            # last letter candidates (keep letters; map 6->G)
            last_candidates = []
            if raw_last == "6":
                last_candidates = ["G"]
            elif raw_last in {"1", "I", "L"}:
                # Sometimes OCR puts 1 instead of I/L for the final letter.
                last_candidates = ["I", "L"]
            else:
                last_candidates = [raw_last]

            for last in last_candidates:
                cand = f"{num}{last}"
                if self._is_valid_spanish_dni(cand):
                    return cand

        return "NOT FOUND"

    # ------------------------
    # Name extraction (unchanged)
    # ------------------------

    def _extract_name_from_mrz(self) -> str:
        """
        MRZ suele venir con '<', ej:
        DEL<CUETO<DEL<POZO<<MANEC
        """
        if not self.text:
            return "NOT FOUND"

        t = (self.text or "").upper()
        mrz_lines = re.findall(r"[A-Z0-9<]{20,}", t)
        if not mrz_lines:
            return "NOT FOUND"

        candidates = [ln for ln in mrz_lines if "<<" in ln and re.match(r"^[A-Z]", ln)]
        if not candidates:
            candidates = [ln for ln in mrz_lines if "<<" in ln]

        if not candidates:
            return "NOT FOUND"

        line = max(candidates, key=len)

        parts = line.split("<<", 1)
        surname = parts[0].replace("<", " ").strip()
        given = parts[1].replace("<", " ").strip() if len(parts) > 1 else ""
        full = re.sub(r"\s+", " ", f"{surname} {given}").strip()

        return full if len(full) >= 6 else "NOT FOUND"

    def _extract_name(self) -> str:
        if not self.text:
            return "NOT FOUND"

        mrz_name = self._extract_name_from_mrz()
        if mrz_name != "NOT FOUND":
            return mrz_name

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

            return name if not self._is_garbage_name(name) else "NOT FOUND"

        lines = [self._clean_name(ln) for ln in t.splitlines() if ln.strip()]
        caps_lines = [ln for ln in lines if re.fullmatch(r"[A-ZÑÁÉÍÓÚ\s]{10,}", ln)]
        if caps_lines:
            cand = max(caps_lines, key=len)
            return cand if not self._is_garbage_name(cand) else "NOT FOUND"

        return "NOT FOUND"

    def _clean_name(self, name: str) -> str:
        return re.sub(r"\s+", " ", (name or "")).strip()

    def _is_garbage_name(self, name: str) -> bool:
        if not name or name == "NOT FOUND":
            return True
        up = name.upper()

        if re.search(r"\b(APELLIDOS|NOMBRE|DOCUMENTO|NIF|DNI|SEXO|NACIONALIDAD|FECHA|CADUCIDAD)\b", up):
            return True

        if len(up) < 10 or len(up.split()) < 2:
            return True

        tokens = up.split()
        short_tokens = sum(1 for tok in tokens if len(tok) <= 2)
        if short_tokens >= 2:
            return True

        return False
