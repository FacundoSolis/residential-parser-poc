"""
Parser for DNI documents

✅ Robust against OCR noise:
- misread digits (O->0, I/L->1)
- last letter misread (G->6)
- MRZ parsing fallback (lines with '<')
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
        t = (t or "").upper()
        t = t.replace(" ", "").replace("-", "").replace(".", "").replace(":", "")
        t = t.replace("O", "0").replace("I", "1").replace("L", "1")
        return t

    def _is_valid_spanish_dni(self, dni: str) -> bool:
        dni = (dni or "").strip().upper().replace(" ", "")
        if not re.match(r"^\d{8}[A-Z]$", dni):
            return False
        letters = "TRWAGMYFPDXBNJZSQVHLCKE"
        num = int(dni[:8])
        return dni[-1] == letters[num % 23]

    def _extract_dni_number(self) -> str:
        """Extract DNI/NIF number (8 digits + letter). Accepts OCR '6' for 'G'."""
        if not self.text:
            return "NOT FOUND"

        t = self._normalize_digits(self.text)

        # Busca TODOS los candidatos y devuelve el primero válido
        for m in re.finditer(r"(\d{8}[A-Z6])", t):
            cand = m.group(1)

            # Caso OCR típico: última letra "G" -> "6"
            if cand[-1] == "6":
                cand_g = cand[:-1] + "G"
                if self._is_valid_spanish_dni(cand_g):
                    return cand_g

            if self._is_valid_spanish_dni(cand):
                return cand

        return "NOT FOUND"


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

        # Preferimos la línea de nombre: empieza por LETRA y contiene '<<'
        candidates = [ln for ln in mrz_lines if "<<" in ln and re.match(r"^[A-Z]", ln)]
        if not candidates:
            candidates = [ln for ln in mrz_lines if "<<" in ln]

        if not candidates:
            return "NOT FOUND"

        # normalmente la de nombre es la más larga/“limpia”
        line = max(candidates, key=len)

        parts = line.split("<<", 1)
        surname = parts[0].replace("<", " ").strip()
        given = parts[1].replace("<", " ").strip() if len(parts) > 1 else ""
        full = re.sub(r"\s+", " ", f"{surname} {given}").strip()

        return full if len(full) >= 6 else "NOT FOUND"


    def _extract_name(self) -> str:
        if not self.text:
            return "NOT FOUND"

        # 1) MRZ primero (en tus OCR suele ser lo único "limpio")
        mrz_name = self._extract_name_from_mrz()
        if mrz_name != "NOT FOUND":
            # el MRZ a veces trae el nombre truncado, pero al menos apellidos salen bien
            return mrz_name

        # 2) fallback a tus heurísticas
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

        # Line heuristics
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
